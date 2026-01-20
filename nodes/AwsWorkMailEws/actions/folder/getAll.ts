import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function getAll(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const parentFolderId = this.getNodeParameter('parentFolderId', index, 'msgfolderroot') as string;

	const results = await ewsClient.getFolders(parentFolderId);

	return results.map(result => ({ json: result as IDataObject }));
}
