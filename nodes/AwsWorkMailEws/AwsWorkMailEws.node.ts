import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';

import { messageOperations, messageFields } from './descriptions/MessageDescription';
import { folderOperations, folderFields } from './descriptions/FolderDescription';
import { calendarOperations, calendarFields } from './descriptions/CalendarDescription';
import { eventOperations, eventFields } from './descriptions/EventDescription';
import { contactOperations, contactFields } from './descriptions/ContactDescription';
import { attachmentOperations, attachmentFields } from './descriptions/AttachmentDescription';

import * as message from './actions/message';
import * as folder from './actions/folder';
import * as calendar from './actions/calendar';
import * as event from './actions/event';
import * as contact from './actions/contact';
import * as attachment from './actions/attachment';

export class AwsWorkMailEws implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'AWS WorkMail (EWS)',
		name: 'awsWorkMailEws',
		icon: 'file:awsworkmail.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Interact with AWS WorkMail using Exchange Web Services (EWS)',
		defaults: {
			name: 'AWS WorkMail (EWS)',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'awsWorkMailEwsApi',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Attachment',
						value: 'attachment',
					},
					{
						name: 'Calendar',
						value: 'calendar',
					},
					{
						name: 'Contact',
						value: 'contact',
					},
					{
						name: 'Event',
						value: 'event',
					},
					{
						name: 'Folder',
						value: 'folder',
					},
					{
						name: 'Message',
						value: 'message',
					},
				],
				default: 'message',
			},
			...messageOperations,
			...messageFields,
			...folderOperations,
			...folderFields,
			...calendarOperations,
			...calendarFields,
			...eventOperations,
			...eventFields,
			...contactOperations,
			...contactFields,
			...attachmentOperations,
			...attachmentFields,
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const resource = this.getNodeParameter('resource', 0);
		const operation = this.getNodeParameter('operation', 0);

		for (let i = 0; i < items.length; i++) {
			try {
				let responseData: INodeExecutionData[];

				if (resource === 'message') {
					if (operation === 'send') {
						responseData = await message.send.call(this, i);
					} else if (operation === 'get') {
						responseData = await message.get.call(this, i);
					} else if (operation === 'getAll') {
						responseData = await message.getAll.call(this, i);
					} else if (operation === 'update') {
						responseData = await message.update.call(this, i);
					} else if (operation === 'delete') {
						responseData = await message.delete.call(this, i);
					} else if (operation === 'move') {
						responseData = await message.move.call(this, i);
					} else if (operation === 'reply') {
						responseData = await message.reply.call(this, i);
					} else {
						throw new Error(`Unknown operation: ${operation}`);
					}
				} else if (resource === 'folder') {
					if (operation === 'create') {
						responseData = await folder.create.call(this, i);
					} else if (operation === 'get') {
						responseData = await folder.get.call(this, i);
					} else if (operation === 'getAll') {
						responseData = await folder.getAll.call(this, i);
					} else if (operation === 'update') {
						responseData = await folder.update.call(this, i);
					} else if (operation === 'delete') {
						responseData = await folder.delete.call(this, i);
					} else if (operation === 'getFolderMessages') {
						responseData = await folder.getFolderMessages.call(this, i);
					} else {
						throw new Error(`Unknown operation: ${operation}`);
					}
				} else if (resource === 'calendar') {
					if (operation === 'create') {
						responseData = await calendar.create.call(this, i);
					} else if (operation === 'get') {
						responseData = await calendar.get.call(this, i);
					} else if (operation === 'getAll') {
						responseData = await calendar.getAll.call(this, i);
					} else if (operation === 'update') {
						responseData = await calendar.update.call(this, i);
					} else if (operation === 'delete') {
						responseData = await calendar.delete.call(this, i);
					} else {
						throw new Error(`Unknown operation: ${operation}`);
					}
				} else if (resource === 'event') {
					if (operation === 'create') {
						responseData = await event.create.call(this, i);
					} else if (operation === 'get') {
						responseData = await event.get.call(this, i);
					} else if (operation === 'getAll') {
						responseData = await event.getAll.call(this, i);
					} else if (operation === 'update') {
						responseData = await event.update.call(this, i);
					} else if (operation === 'delete') {
						responseData = await event.delete.call(this, i);
					} else {
						throw new Error(`Unknown operation: ${operation}`);
					}
				} else if (resource === 'contact') {
					if (operation === 'create') {
						responseData = await contact.create.call(this, i);
					} else if (operation === 'get') {
						responseData = await contact.get.call(this, i);
					} else if (operation === 'getAll') {
						responseData = await contact.getAll.call(this, i);
					} else if (operation === 'update') {
						responseData = await contact.update.call(this, i);
					} else if (operation === 'delete') {
						responseData = await contact.delete.call(this, i);
					} else {
						throw new Error(`Unknown operation: ${operation}`);
					}
				} else if (resource === 'attachment') {
					if (operation === 'add') {
						responseData = await attachment.add.call(this, i);
					} else if (operation === 'get') {
						responseData = await attachment.get.call(this, i);
					} else if (operation === 'getAll') {
						responseData = await attachment.getAll.call(this, i);
					} else if (operation === 'download') {
						responseData = await attachment.download.call(this, i);
					} else {
						throw new Error(`Unknown operation: ${operation}`);
					}
				} else {
					throw new Error(`Unknown resource: ${resource}`);
				}

				returnData.push(...responseData);
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({
						json: {
							error: (error as Error).message || 'Unknown error',
						},
						pairedItem: {
							item: i,
						},
					});
					continue;
				}
				throw error;
			}
		}

		return [returnData];
	}
}
