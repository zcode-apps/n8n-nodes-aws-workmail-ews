import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function get(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const calendarId = this.getNodeParameter('calendarId', index) as string;

	const result = await ewsClient.getCalendar(calendarId);

	return [{ json: result as IDataObject }];
}
