import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { z } from 'zod';
import { EwsClient } from '../../transport/EwsClient';

// Sicherheits-Schemas für Input-Validierung
const emailSchema = z.string().email().max(254).regex(/^[^<>&"']*$/, 'Email darf keine HTML-Tags enthalten');
const emailListSchema = z.string().max(1000).transform((val) =>
  val.split(',').map(email => email.trim()).filter(email => email.length > 0)
);
const subjectSchema = z.string().max(255).regex(/^[^<>&]*$/, 'Betreff darf keine HTML-Tags enthalten');
const bodySchema = z.string().max(100000); // 100KB Limit für Email-Body

export async function send(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const toRecipientsRaw = this.getNodeParameter('toRecipients', index) as string;
	const subjectRaw = this.getNodeParameter('subject', index) as string;
	const bodyRaw = this.getNodeParameter('body', index) as string;
	const bodyType = this.getNodeParameter('bodyType', index, 'HTML') as string;
	const additionalFields = this.getNodeParameter('additionalFields', index, {}) as IDataObject;

	// Sicherheitsvalidierung der Eingaben
	const validatedRecipients = emailListSchema.parse(toRecipientsRaw);
	const validatedSubject = subjectSchema.parse(subjectRaw);
	const validatedBody = bodySchema.parse(bodyRaw);

	// Zusätzliche E-Mail-Validierung für jeden Empfänger
	const validEmails: string[] = [];
	const invalidEmails: string[] = [];

	for (const email of validatedRecipients) {
		const result = emailSchema.safeParse(email);
		if (result.success) {
			validEmails.push(result.data);
		} else {
			invalidEmails.push(email);
		}
	}

	if (invalidEmails.length > 0) {
		throw new Error(`Ungültige E-Mail-Adressen: ${invalidEmails.join(', ')}`);
	}

	if (validEmails.length === 0) {
		throw new Error('Mindestens eine gültige E-Mail-Adresse muss angegeben werden');
	}

	const messageData: IDataObject = {
		subject: validatedSubject,
		body: validatedBody,
		bodyType,
		toRecipients: validEmails,
	};

	if (additionalFields.ccRecipients) {
		const ccRecipientsRaw = additionalFields.ccRecipients as string;
		const validatedCC = emailListSchema.parse(ccRecipientsRaw);
		const validCC: string[] = [];

		for (const email of validatedCC) {
			const result = emailSchema.safeParse(email);
			if (result.success) {
				validCC.push(result.data);
			} else {
				invalidEmails.push(`CC: ${email}`);
			}
		}

		if (validCC.length > 0) {
			messageData.ccRecipients = validCC;
		}
	}

	if (additionalFields.bccRecipients) {
		const bccRecipientsRaw = additionalFields.bccRecipients as string;
		const validatedBCC = emailListSchema.parse(bccRecipientsRaw);
		const validBCC: string[] = [];

		for (const email of validatedBCC) {
			const result = emailSchema.safeParse(email);
			if (result.success) {
				validBCC.push(result.data);
			} else {
				invalidEmails.push(`BCC: ${email}`);
			}
		}

		if (validBCC.length > 0) {
			messageData.bccRecipients = validBCC;
		}
	}

	if (additionalFields.importance) {
		messageData.importance = additionalFields.importance;
	}

	if (additionalFields.sensitivity) {
		messageData.sensitivity = additionalFields.sensitivity;
	}

	if (additionalFields.categories) {
		messageData.categories = (additionalFields.categories as string).split(',').map(cat => cat.trim());
	}

	const result = await ewsClient.sendMessage(messageData);

	return [{ json: result as IDataObject }];
}
