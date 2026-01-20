import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function getAll(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const folderId = this.getNodeParameter('folderId', index, 'inbox') as string;
	const returnAll = this.getNodeParameter('returnAll', index, false) as boolean;
	const additionalFields = this.getNodeParameter('additionalFields', index, {}) as IDataObject;

	const options: any = {};

	if (!returnAll) {
		const limit = this.getNodeParameter('limit', index, 50) as number;
		options.maxResults = limit;
	} else {
		options.maxResults = 500; // Maximum allowed by EWS
	}

	if (additionalFields.sort) {
		options.sort = additionalFields.sort;
	}

	const results = await ewsClient.getMessages(folderId, options);

	return results.map(result => ({ json: result as IDataObject }));
}
