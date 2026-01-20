import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { z } from 'zod';
import { EwsClient } from '../../transport/EwsClient';

// Sicherheits-Schema für Reply-Draft-Body
const replyBodySchema = z.string().max(50000); // 50KB Limit für Reply-Body

export async function createReplyDraft(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageId = this.getNodeParameter('messageId', index) as string;
	const replyBodyRaw = this.getNodeParameter('draftReplyBody', index) as string;
	const bodyType = this.getNodeParameter('draftBodyType', index, 'HTML') as 'HTML' | 'Text';
	const replyAll = this.getNodeParameter('draftReplyAll', index, false) as boolean;

	// Sicherheitsvalidierung
	const replyBody = replyBodySchema.parse(replyBodyRaw);

	const result = await ewsClient.createReplyDraft(messageId, replyBody, replyAll, bodyType);

	return [{ json: result as IDataObject }];
}
