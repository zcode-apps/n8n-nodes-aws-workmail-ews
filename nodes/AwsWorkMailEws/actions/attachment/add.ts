import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function add(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const messageId = this.getNodeParameter('messageId', index) as string;
	const binaryPropertyName = this.getNodeParameter('binaryPropertyName', index, 'data') as string;

	const binaryData = this.helpers.assertBinaryData(index, binaryPropertyName);

	let fileContent: string;

	if (binaryData.id) {
		const binaryDataBuffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
		fileContent = binaryDataBuffer.toString('base64');
	} else {
		fileContent = binaryData.data;
	}

	const attachmentData: IDataObject = {
		name: binaryData.fileName || 'attachment',
		content: fileContent,
		contentType: binaryData.mimeType || 'application/octet-stream',
	};

	const result = await ewsClient.addAttachment(messageId, attachmentData);

	return [{ json: result as IDataObject }];
}
