import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function get(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const attachmentId = this.getNodeParameter('attachmentId', index) as string;

	const result = await ewsClient.getAttachment(attachmentId);

	return [{ json: result as IDataObject }];
}
