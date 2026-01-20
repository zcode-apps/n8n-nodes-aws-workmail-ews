import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

export async function create(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const calendarId = this.getNodeParameter('calendarId', index, 'calendar') as string;
	const subject = this.getNodeParameter('subject', index) as string;
	const start = this.getNodeParameter('start', index) as string;
	const end = this.getNodeParameter('end', index) as string;
	const additionalFields = this.getNodeParameter('additionalFields', index, {}) as IDataObject;

	const eventData: IDataObject = {
		subject,
		start,
		end,
	};

	if (additionalFields.body) {
		eventData.body = additionalFields.body;
		eventData.bodyType = additionalFields.bodyType || 'HTML';
	}

	if (additionalFields.location) {
		eventData.location = additionalFields.location;
	}

	if (additionalFields.isAllDayEvent !== undefined) {
		eventData.isAllDayEvent = additionalFields.isAllDayEvent;
	}

	if (additionalFields.requiredAttendees) {
		eventData.requiredAttendees = (additionalFields.requiredAttendees as string).split(',').map(email => email.trim());
	}

	if (additionalFields.optionalAttendees) {
		eventData.optionalAttendees = (additionalFields.optionalAttendees as string).split(',').map(email => email.trim());
	}

	if (additionalFields.importance) {
		eventData.importance = additionalFields.importance;
	}

	if (additionalFields.sensitivity) {
		eventData.sensitivity = additionalFields.sensitivity;
	}

	if (additionalFields.categories) {
		eventData.categories = (additionalFields.categories as string).split(',').map(cat => cat.trim());
	}

	const result = await ewsClient.createEvent(calendarId, eventData);

	return [{ json: result as IDataObject }];
}
