import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function get(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const eventId = this.getNodeParameter('eventId', index) as string;

	const result = await ewsClient.getEvent(eventId);

	return [{ json: result as IDataObject }];
}
