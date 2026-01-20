import { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function del(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const contactId = this.getNodeParameter('contactId', index) as string;

	await ewsClient.deleteContact(contactId);

	return [{ json: { success: true, contactId } }];
}
