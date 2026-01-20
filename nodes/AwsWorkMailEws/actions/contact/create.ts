import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function create(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const displayName = this.getNodeParameter('displayName', index) as string;
	const emailAddress = this.getNodeParameter('emailAddress', index) as string;
	const additionalFields = this.getNodeParameter('additionalFields', index, {}) as IDataObject;

	const contactData: IDataObject = {
		displayName,
		emailAddress,
	};

	if (additionalFields.givenName) {
		contactData.givenName = additionalFields.givenName;
	}

	if (additionalFields.surname) {
		contactData.surname = additionalFields.surname;
	}

	if (additionalFields.companyName) {
		contactData.companyName = additionalFields.companyName;
	}

	if (additionalFields.jobTitle) {
		contactData.jobTitle = additionalFields.jobTitle;
	}

	if (additionalFields.phoneNumber) {
		contactData.phoneNumber = additionalFields.phoneNumber;
	}

	if (additionalFields.categories) {
		contactData.categories = (additionalFields.categories as string).split(',').map(cat => cat.trim());
	}

	const result = await ewsClient.createContact(contactData);

	return [{ json: result as IDataObject }];
}
