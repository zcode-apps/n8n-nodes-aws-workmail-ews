import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { z } from 'zod';
import { EwsClient } from '../../transport/EwsClient';

// Sicherheits-Schemas für Update-Operationen
const messageIdSchema = z.string().max(1024).regex(/^[A-Za-z0-9+/=_-]+$/, 'Ungültige Message-ID');
const updateFieldsSchema = z.object({
	IsRead: z.boolean().optional(),
	Subject: z.string().max(255).regex(/^[^<>&]*$/, 'Betreff darf keine HTML-Tags enthalten').optional(),
}).strict(); // Nur erlaubte Felder

export async function update(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageIdRaw = this.getNodeParameter('messageId', index) as string;
	const updateFieldsRaw = this.getNodeParameter('updateFields', index, {}) as IDataObject;

	// Sicherheitsvalidierung
	const messageId = messageIdSchema.parse(messageIdRaw);
	const updateFields = updateFieldsSchema.parse(updateFieldsRaw);

	const result = await ewsClient.updateMessage(messageId, updateFields);

	return [{ json: result as IDataObject }];
}
