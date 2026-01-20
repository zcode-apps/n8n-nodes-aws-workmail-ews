import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function get(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageId = this.getNodeParameter('messageId', index) as string;

	const result = await ewsClient.getMessage(messageId);

	return [{ json: result as IDataObject }];
}
