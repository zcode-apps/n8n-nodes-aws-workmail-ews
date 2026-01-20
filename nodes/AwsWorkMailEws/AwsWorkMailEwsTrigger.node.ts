import {
	IDataObject,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IPollFunctions,
} from 'n8n-workflow';

import { EwsClient } from './transport/EwsClient';

export class AwsWorkMailEwsTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'AWS WorkMail Trigger (EWS)',
		name: 'awsWorkMailEwsTrigger',
		icon: 'file:awsworkmail.svg',
		group: ['trigger'],
		version: 1,
		subtitle: 'Poll for new messages',
		description: 'Starts workflow when a new email is received in AWS WorkMail',
		defaults: {
			name: 'AWS WorkMail Trigger (EWS)',
		},
		inputs: [],
		outputs: ['main'],
		credentials: [
			{
				name: 'awsWorkMailEwsApi',
				required: true,
			},
		],
		polling: true,
		properties: [
			{
				displayName: 'Folder',
				name: 'folderId',
				type: 'string',
				default: 'inbox',
				description: 'Folder ID or distinguished folder name to monitor (inbox, sentitems, drafts, etc.)',
				required: true,
			},
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				options: [
					{
						displayName: 'Max Items',
						name: 'maxItems',
						type: 'number',
						typeOptions: {
							minValue: 1,
							maxValue: 100,
						},
						default: 10,
						description: 'Maximum number of messages to fetch per poll',
					},
					{
						displayName: 'Download Attachments',
						name: 'downloadAttachments',
						type: 'boolean',
						default: false,
						description: 'Whether to download attachments as binary data',
					},
				],
			},
		],
	};

	async poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
		// Use 'workflow' scope for static data - more reliable across n8n versions
		const webhookData = this.getWorkflowStaticData('node');
		const folderId = this.getNodeParameter('folderId', 'inbox') as string;
		const options = this.getNodeParameter('options', {}) as IDataObject;
		const maxItems = (options.maxItems as number) || 10;
		const downloadAttachments = (options.downloadAttachments as boolean) || false;

		const credentials = await this.getCredentials('awsWorkMailEwsApi');
		const ewsClient = new EwsClient(credentials as any);

		// Determine if this is a manual test execution or automatic polling
		const mode = this.getMode();
		const isManualTrigger = mode === 'manual';

		// Fetch messages from EWS
		let messages: any[];
		try {
			messages = await ewsClient.getMessages(folderId, {
				maxResults: Math.max(maxItems * 3, 50),
			});
		} catch (error) {
			// If there's an error fetching messages, log it and return null
			console.error('AWS WorkMail Trigger: Error fetching messages:', error);
			throw error;
		}

		// Sort messages by DateTimeReceived (newest first)
		messages.sort((a: any, b: any) => {
			const dateA = a.DateTimeReceived ? new Date(a.DateTimeReceived).getTime() : 0;
			const dateB = b.DateTimeReceived ? new Date(b.DateTimeReceived).getTime() : 0;
			return dateB - dateA; // DESC order
		});

		// Helper function to process messages
		const processMessages = async (messagesToProcess: any[]): Promise<INodeExecutionData[]> => {
			const returnData: INodeExecutionData[] = [];

			for (const message of messagesToProcess) {
				const item: INodeExecutionData = {
					json: message as IDataObject,
				};

				// Download attachments if requested
				if (downloadAttachments && message.HasAttachments && message.ItemId?.Id) {
					try {
						const attachments = await ewsClient.getAttachments(message.ItemId.Id);

						if (attachments.length > 0) {
							item.binary = {};

							for (let i = 0; i < attachments.length; i++) {
								const att = attachments[i];
								if (att && att.AttachmentId?.Id) {
									const buffer = await ewsClient.downloadAttachment(att.AttachmentId.Id);

									const binaryData = await this.helpers.prepareBinaryData(
										buffer,
										(att.Name as string) || `attachment_${i}`,
										(att.ContentType as string) || 'application/octet-stream',
									);

									item.binary[`attachment_${i}`] = binaryData;
								}
							}
						}
					} catch (error) {
						console.error('Error downloading attachments:', error);
					}
				}

				returnData.push(item);
			}

			return returnData;
		};

		// For MANUAL TEST: Return the latest messages without filtering
		if (isManualTrigger) {
			const limitedMessages = messages.slice(0, maxItems);
			
			if (limitedMessages.length === 0) {
				return null;
			}

			return [await processMessages(limitedMessages)];
		}

		// =====================================================================
		// AUTOMATIC POLLING: Use message ID tracking (most reliable method)
		// =====================================================================
		
		// Get all current message IDs from the inbox
		const currentMessageIds = messages
			.filter((m: any) => m.ItemId?.Id)
			.map((m: any) => m.ItemId.Id as string);

		// Get stored data from previous polls
		const knownMessageIds: string[] = (webhookData.knownMessageIds as string[]) || [];
		const isFirstPoll = knownMessageIds.length === 0 && !webhookData.initialized;

		if (isFirstPoll) {
			// FIRST POLL: Store all current message IDs
			// Mark all existing messages as "known" so we don't trigger for them
			webhookData.initialized = true;
			webhookData.knownMessageIds = currentMessageIds.slice(0, 500);
			
			// Return null - don't trigger for existing emails
			return null;
		}

		// Find NEW messages: IDs in current list that are NOT in known list
		const newMessageIds = currentMessageIds.filter(id => !knownMessageIds.includes(id));
		
		if (newMessageIds.length === 0) {
			// No new messages - nothing to do
			return null;
		}

		// Get the actual message objects for the new IDs
		const newMessages = messages.filter((m: any) => 
			m.ItemId?.Id && newMessageIds.includes(m.ItemId.Id)
		);

		// Limit to maxItems
		const limitedMessages = newMessages.slice(0, maxItems);

		if (limitedMessages.length > 0) {
			// Update known message IDs - add all current IDs to prevent future duplicates
			// Keep a rolling window of the last 500 IDs
			webhookData.knownMessageIds = currentMessageIds.slice(0, 500);
			
			return [await processMessages(limitedMessages)];
		}

		return null;
	}
}
