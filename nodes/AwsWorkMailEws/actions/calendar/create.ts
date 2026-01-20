import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function create(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const calendarName = this.getNodeParameter('calendarName', index) as string;

	const result = await ewsClient.createCalendar(calendarName);

	return [{ json: result as IDataObject }];
}
