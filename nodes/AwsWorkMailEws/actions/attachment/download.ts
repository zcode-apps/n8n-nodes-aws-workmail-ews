import { IExecuteFunctions, IDataObject, INodeExecutionData } from 'n8n-workflow';
import { EwsClient } from '../../transport/EwsClient';

// Sicherheitskonstanten für Attachments
const MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024; // 25MB Limit
const ALLOWED_MIME_TYPES = [
	'application/pdf',
	'image/jpeg',
	'image/png',
	'image/gif',
	'image/webp',
	'application/msword',
	'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	'application/vnd.ms-excel',
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	'text/plain',
	'text/csv',
	'application/zip',
];

export async function download(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const credentials = await this.getCredentials('awsWorkMailEwsApi');
	const ewsClient = new EwsClient(credentials as any);

	const attachmentId = this.getNodeParameter('attachmentId', index) as string;
	const binaryPropertyName = this.getNodeParameter('binaryPropertyName', index, 'data') as string;

	// Validierung der Attachment-ID (sollte eine sichere Zeichenkette sein)
	if (!attachmentId || typeof attachmentId !== 'string' || attachmentId.length > 200) {
		throw new Error('Ungültige Attachment-ID');
	}

	const attachment = await ewsClient.getAttachment(attachmentId);

	// Sicherheitsprüfungen vor dem Download
	if (attachment.Size && attachment.Size > MAX_ATTACHMENT_SIZE) {
		throw new Error(`Attachment zu groß: ${(attachment.Size / 1024 / 1024).toFixed(2)}MB. Maximum erlaubt: ${(MAX_ATTACHMENT_SIZE / 1024 / 1024).toFixed(0)}MB`);
	}

	if (attachment.ContentType && !ALLOWED_MIME_TYPES.includes(attachment.ContentType.toLowerCase())) {
		throw new Error(`Nicht erlaubter Dateityp: ${attachment.ContentType}`);
	}

	// Sicherheitsprüfung für Dateinamen
	if (attachment.Name) {
		const fileName = attachment.Name as string;
		// Prüfe auf gefährliche Dateinamen (Path Traversal)
		if (fileName.includes('..') || fileName.includes('/') || fileName.includes('\\') || fileName.startsWith('.')) {
			throw new Error('Unsicherer Dateiname erkannt');
		}
		// Prüfe auf verdächtige Dateierweiterungen
		const dangerousExtensions = ['.exe', '.bat', '.cmd', '.scr', '.pif', '.com', '.jar', '.js', '.vbs', '.wsf'];
		const extension = fileName.toLowerCase().substring(fileName.lastIndexOf('.'));
		if (dangerousExtensions.includes(extension)) {
			throw new Error(`Nicht erlaubte Dateierweiterung: ${extension}`);
		}
	}

	const buffer = await ewsClient.downloadAttachment(attachmentId);

	const binaryData = await this.helpers.prepareBinaryData(
		buffer,
		attachment.Name as string || 'attachment',
		attachment.ContentType as string || 'application/octet-stream',
	);

	return [{
		json: attachment as IDataObject,
		binary: {
			[binaryPropertyName]: binaryData,
		},
	}];
}
