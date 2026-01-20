import { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function del(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const folderId = this.getNodeParameter('folderId', index) as string;

	await ewsClient.deleteFolder(folderId, false);

	return [{ json: { success: true, folderId } }];
}
