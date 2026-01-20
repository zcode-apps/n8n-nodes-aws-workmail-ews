import { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function del(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageId = this.getNodeParameter('messageId', index) as string;

	await ewsClient.deleteMessage(messageId, true);

	return [{ json: { success: true, messageId } }];
}
