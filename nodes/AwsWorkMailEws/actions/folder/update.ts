import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function update(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const folderId = this.getNodeParameter('folderId', index) as string;
	const updateFields = this.getNodeParameter('updateFields', index, {}) as IDataObject;

	const result = await ewsClient.updateFolder(folderId, updateFields);

	return [{ json: result as IDataObject }];
}
