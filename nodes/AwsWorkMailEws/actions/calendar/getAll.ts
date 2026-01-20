import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function getAll(
	this: IExecuteFunctions,
	_index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const results = await ewsClient.getCalendars();

	return results.map(result => ({ json: result as IDataObject }));
}
