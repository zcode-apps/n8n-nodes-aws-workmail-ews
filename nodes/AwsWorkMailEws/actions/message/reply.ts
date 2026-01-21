import { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { z } from 'zod';
import { EwsClient } from '../../transport/EwsClient';

// Sicherheits-Schema für Reply-Body
const replyBodySchema = z.string().max(50000); // 50KB Limit für Reply-Body
const messageIdSchema = z.string().max(1024).regex(/^[A-Za-z0-9+/=_-]+$/, 'Ungültige Message-ID');

export async function reply(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageIdRaw = this.getNodeParameter('messageId', index) as string;
	const replyBodyRaw = this.getNodeParameter('replyBody', index) as string;
	const replyAll = this.getNodeParameter('replyAll', index, false) as boolean;

	// Sicherheitsvalidierung
	const messageId = messageIdSchema.parse(messageIdRaw);
	const replyBody = replyBodySchema.parse(replyBodyRaw);

	await ewsClient.replyToMessage(messageId, replyBody, replyAll);

	return [{ json: { success: true, messageId, replySent: true } }];
}
