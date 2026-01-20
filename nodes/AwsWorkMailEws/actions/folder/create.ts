import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function create(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const folderName = this.getNodeParameter('folderName', index) as string;
	const parentFolderId = this.getNodeParameter('parentFolderId', index, 'msgfolderroot') as string;

	const result = await ewsClient.createFolder(folderName, parentFolderId);

	return [{ json: result as IDataObject }];
}
