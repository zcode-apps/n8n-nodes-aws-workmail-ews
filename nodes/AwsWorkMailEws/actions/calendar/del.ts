import { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function del(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const calendarId = this.getNodeParameter('calendarId', index) as string;

	await ewsClient.deleteCalendar(calendarId);

	return [{ json: { success: true, calendarId } }];
}
