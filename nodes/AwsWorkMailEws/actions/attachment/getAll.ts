import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function getAll(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageId = this.getNodeParameter('messageId', index) as string;

	const results = await ewsClient.getAttachments(messageId);

	return results.map(result => ({ json: result as IDataObject }));
}
