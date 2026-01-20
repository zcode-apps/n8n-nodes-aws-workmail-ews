import { IDataObject, NodeApiError } from 'n8n-workflow';
import * as ews from 'ews-javascript-api';
import {
	IEwsCredentials,
	IEwsMessage,
	IEwsFolder,
	IEwsCalendar,
	IEwsCalendarItem,
	IEwsContact,
	IEwsAttachment,
	IEwsSearchOptions,
} from './types';

export class EwsClient {
	private service: ews.ExchangeService;

	constructor(credentials: IEwsCredentials) {
		// Create ExchangeService
		this.service = new ews.ExchangeService(ews.ExchangeVersion.Exchange2010_SP2);

		// Set credentials
		this.service.Credentials = new ews.WebCredentials(
			credentials.username,
			credentials.password
		);

		// Set URL
		this.service.Url = new ews.Uri(credentials.ewsUrl);
	}

	private handleError(error: any, operation: string): never {
		// Extrahiere Fehlermeldung aus verschiedenen moeglichen Quellen
		let safeMessage = 'Unbekannter Fehler';

		// Versuche verschiedene Fehlerquellen
		if (error.message) {
			safeMessage = error.message;
		} else if (error.Message) {
			safeMessage = error.Message;
		} else if (error.ErrorMessage) {
			safeMessage = error.ErrorMessage;
		} else if (typeof error === 'string') {
			safeMessage = error;
		}
		
		// EWS-spezifische Fehlerdetails extrahieren
		if (error.Response?.ErrorMessage) {
			safeMessage += ` - ${error.Response.ErrorMessage}`;
		}
		if (error.ResponseCode) {
			safeMessage += ` (Code: ${error.ResponseCode})`;
		}

		// Entferne nur wirklich sensible Daten (Passwoerter, Tokens)
		safeMessage = safeMessage
			.replace(/password[=:]\s*\S+/gi, 'password=[REDACTED]')
			.replace(/token[=:]\s*\S+/gi, 'token=[REDACTED]')
			.replace(/key[=:]\s*\S+/gi, 'key=[REDACTED]')
			.substring(0, 500); // Erlaube laengere Meldungen

		// Beschreibung mit mehr Details
		let description = safeMessage;
		if (error.stack && process.env.NODE_ENV === 'development') {
			description = `${safeMessage} | Stack: ${error.stack.substring(0, 300)}`;
		}

		throw new NodeApiError(
			{ message: `EWS ${operation} fehlgeschlagen` } as any,
			{ message: safeMessage, description },
		);
	}

	// ===============================
	// MESSAGE OPERATIONS
	// ===============================

	async sendMessage(message: IDataObject): Promise<IEwsMessage> {
		try {
			const emailMessage = new ews.EmailMessage(this.service);

			// Set recipients
			if (message.toRecipients) {
				const recipients = Array.isArray(message.toRecipients)
					? message.toRecipients
					: [message.toRecipients];
				recipients.forEach((email: any) => {
					emailMessage.ToRecipients.Add(typeof email === 'string' ? email : email.EmailAddress);
				});
			}

			if (message.ccRecipients) {
				const recipients = Array.isArray(message.ccRecipients)
					? message.ccRecipients
					: [message.ccRecipients];
				recipients.forEach((email: any) => {
					emailMessage.CcRecipients.Add(typeof email === 'string' ? email : email.EmailAddress);
				});
			}

			if (message.bccRecipients) {
				const recipients = Array.isArray(message.bccRecipients)
					? message.bccRecipients
					: [message.bccRecipients];
				recipients.forEach((email: any) => {
					emailMessage.BccRecipients.Add(typeof email === 'string' ? email : email.EmailAddress);
				});
			}

			// Set subject and body
			if (message.subject) {
				emailMessage.Subject = message.subject as string;
			}

			if (message.body) {
				const bodyType = message.bodyType === 'Text'
					? ews.BodyType.Text
					: ews.BodyType.HTML;
				emailMessage.Body = new ews.MessageBody(bodyType, message.body as string);
			}

			// Set importance
			if (message.importance) {
				emailMessage.Importance = this.mapImportance(message.importance as string);
			}

			// Set sensitivity
			if (message.sensitivity) {
				emailMessage.Sensitivity = this.mapSensitivity(message.sensitivity as string);
			}

			// Send and save
			await emailMessage.SendAndSaveCopy();

			return {
				Subject: emailMessage.Subject,
				success: true,
			};
		} catch (error) {
			this.handleError(error, 'SendMessage');
		}
	}

	async getMessage(messageId: string): Promise<IEwsMessage> {
		try {
			const itemId = new ews.ItemId(messageId);
			const propertySet = new ews.PropertySet(ews.BasePropertySet.FirstClassProperties, [ews.ItemSchema.Body]);
			propertySet.RequestedBodyType = ews.BodyType.Text;
			
			const message = await ews.EmailMessage.Bind(
				this.service,
				itemId,
				propertySet
			);

			return this.convertMessageToJson(message);
		} catch (error) {
			this.handleError(error, 'GetMessage');
		}
	}

	async getMessages(folderId = 'inbox', options: IEwsSearchOptions = {}): Promise<IEwsMessage[]> {
		try {
			const maxResults = options.maxResults || 50;
			const offset = options.offset || 0;

			// Get folder ID
			const folderIdObj = this.getFolderId(folderId);

			// Create view
			const view = new ews.ItemView(maxResults, offset);
			view.PropertySet = new ews.PropertySet(ews.BasePropertySet.FirstClassProperties);

			// Find items
			const findResults = await this.service.FindItems(folderIdObj, view);

			// Load full message details (including Body)
			const messages: IEwsMessage[] = [];
			const bodyPropertySet = new ews.PropertySet(ews.BasePropertySet.FirstClassProperties, [ews.ItemSchema.Body]);
			bodyPropertySet.RequestedBodyType = ews.BodyType.Text;
			
			for (const item of findResults.Items) {
				if (item instanceof ews.EmailMessage) {
					// Bind to get full message with Body
					const fullMessage = await ews.EmailMessage.Bind(
						this.service,
						item.Id,
						bodyPropertySet
					);
					messages.push(this.convertMessageToJson(fullMessage));
				}
			}

			return messages;
		} catch (error) {
			this.handleError(error, 'GetMessages');
		}
	}

	async updateMessage(messageId: string, updates: IDataObject): Promise<IEwsMessage> {
		try {
			const itemId = new ews.ItemId(messageId);
			const message = await ews.EmailMessage.Bind(this.service, itemId);

			if (updates.IsRead !== undefined) {
				message.IsRead = updates.IsRead as boolean;
			}

			if (updates.Subject) {
				message.Subject = updates.Subject as string;
			}

			await message.Update(ews.ConflictResolutionMode.AlwaysOverwrite);

			return this.convertMessageToJson(message);
		} catch (error) {
			this.handleError(error, 'UpdateMessage');
		}
	}

	async deleteMessage(messageId: string, moveToDeletedItems = true): Promise<void> {
		try {
			const itemId = new ews.ItemId(messageId);
			const deleteMode = moveToDeletedItems
				? ews.DeleteMode.MoveToDeletedItems
				: ews.DeleteMode.HardDelete;

			const message = await ews.EmailMessage.Bind(this.service, itemId);
			await message.Delete(deleteMode);
		} catch (error) {
			this.handleError(error, 'DeleteMessage');
		}
	}

	async moveMessage(messageId: string, targetFolderId: string): Promise<IEwsMessage> {
		try {
			const itemId = new ews.ItemId(messageId);
			const targetFolder = this.getFolderId(targetFolderId);

			const response = await this.service.MoveItems([itemId], targetFolder);
			if (!response.Responses[0]?.Item?.Id) {
				throw new Error('Move operation failed: No valid response received');
			}
			const message = await ews.EmailMessage.Bind(this.service, response.Responses[0].Item.Id);

			return this.convertMessageToJson(message);
		} catch (error) {
			this.handleError(error, 'MoveMessage');
		}
	}

	async replyToMessage(messageId: string, replyBody: string, replyAll = false): Promise<void> {
		try {
			const itemId = new ews.ItemId(messageId);
			const message = await ews.EmailMessage.Bind(this.service, itemId);

			const reply = message.CreateReply(replyAll);

			reply.BodyPrefix = new ews.MessageBody(ews.BodyType.HTML, replyBody);

			await reply.SendAndSaveCopy();
		} catch (error) {
			this.handleError(error, 'ReplyToMessage');
		}
	}

	async createReplyDraft(messageId: string, replyBody: string, _replyAll = false, bodyType: 'HTML' | 'Text' = 'HTML'): Promise<IEwsMessage> {
		// Retry-Funktion fuer EWS-Operationen
		const retry = async <T>(fn: () => Promise<T>, maxRetries = 2, delay = 1000): Promise<T> => {
			let lastError: any;
			for (let attempt = 0; attempt <= maxRetries; attempt++) {
				try {
					return await fn();
				} catch (error: any) {
					lastError = error;
					if (attempt < maxRetries) {
						// Warte vor dem naechsten Versuch
						await new Promise(resolve => setTimeout(resolve, delay * (attempt + 1)));
					}
				}
			}
			throw lastError;
		};

		// Lade die Original-Nachricht mit Retry
		const itemId = new ews.ItemId(messageId);
		const propertySet = new ews.PropertySet(ews.BasePropertySet.FirstClassProperties);
		
		let originalMessage: ews.EmailMessage;
		try {
			originalMessage = await retry(() => ews.EmailMessage.Bind(this.service, itemId, propertySet));
		} catch (error: any) {
			const errorMsg = error?.message || error?.Response?.ErrorMessage || 'Fehler beim Laden der Original-Nachricht';
			this.handleError({ message: errorMsg, ...error }, 'CreateReplyDraft - Bind');
		}

		// Hole Absender-Info sicher
		let fromAddress = '';
		try {
			if (originalMessage!.From) {
				fromAddress = originalMessage!.From.Address || '';
			}
		} catch {
			// From nicht verfuegbar
		}
		
		// Falls keine From-Adresse, Fehler werfen
		if (!fromAddress) {
			this.handleError({ message: 'Keine Absender-Adresse in der Original-Nachricht gefunden' }, 'CreateReplyDraft');
		}

		// Erstelle eine neue EmailMessage als Entwurf
		const draftMessage = new ews.EmailMessage(this.service);
		
		// Setze den Betreff mit "Re: " Prefix
		const originalSubject = originalMessage!.Subject || '';
		draftMessage.Subject = originalSubject.startsWith('Re: ') ? originalSubject : `Re: ${originalSubject}`;
		
		// Setze den Body
		const ewsBodyType = bodyType === 'Text' ? ews.BodyType.Text : ews.BodyType.HTML;
		draftMessage.Body = new ews.MessageBody(ewsBodyType, replyBody);
		
		// Setze Empfaenger
		draftMessage.ToRecipients.Add(fromAddress);
		
		// Speichere als Entwurf im Drafts-Ordner mit Retry
		try {
			await retry(() => draftMessage.Save(ews.WellKnownFolderName.Drafts));
		} catch (error: any) {
			const errorMsg = error?.message || error?.Response?.ErrorMessage || error?.toString() || 'Fehler beim Speichern des Entwurfs';
			// Extrahiere mehr Details aus dem EWS-Fehler
			let detailedMsg = errorMsg;
			if (error?.Response) {
				try {
					detailedMsg += ` (Response: ${JSON.stringify(error.Response).substring(0, 200)})`;
				} catch {
					// JSON stringify fehlgeschlagen
				}
			}
			if (error?.InnerException) {
				detailedMsg += ` (Inner: ${error.InnerException.message || error.InnerException})`;
			}
			this.handleError({ message: detailedMsg, ResponseCode: error?.ResponseCode }, 'CreateReplyDraft - Save');
		}
		
		// Lade den gespeicherten Entwurf fuer die Rueckgabe
		if (draftMessage.Id) {
			try {
				const savedDraft = await ews.EmailMessage.Bind(
					this.service,
					draftMessage.Id,
					new ews.PropertySet(ews.BasePropertySet.FirstClassProperties)
				);
				return this.convertMessageToJson(savedDraft);
			} catch {
				// Falls Bind fehlschlaegt, trotzdem Erfolg melden
			}
		}

		return {
			Subject: draftMessage.Subject,
			success: true,
			savedAsDraft: true,
			folder: 'Drafts',
			toRecipient: fromAddress,
		};
	}

	// ===============================
	// FOLDER OPERATIONS
	// ===============================

	async createFolder(folderName: string, parentFolderId = 'msgfolderroot'): Promise<IEwsFolder> {
		try {
			const folder = new ews.Folder(this.service);
			folder.DisplayName = folderName;

			const parentFolder = this.getFolderId(parentFolderId);
			await folder.Save(parentFolder);

			return {
				FolderId: { Id: folder.Id.UniqueId },
				DisplayName: folder.DisplayName,
			};
		} catch (error) {
			this.handleError(error, 'CreateFolder');
		}
	}

	async getFolder(folderId: string): Promise<IEwsFolder> {
		try {
			const folderIdObj = this.getFolderId(folderId);
			const folder = await ews.Folder.Bind(this.service, folderIdObj);

			return {
				FolderId: { Id: folder.Id.UniqueId },
				DisplayName: folder.DisplayName,
				TotalCount: folder.TotalCount,
				ChildFolderCount: folder.ChildFolderCount,
				UnreadCount: folder.UnreadCount,
			};
		} catch (error) {
			this.handleError(error, 'GetFolder');
		}
	}

	async getFolders(parentFolderId = 'msgfolderroot'): Promise<IEwsFolder[]> {
		try {
			const parentFolder = this.getFolderId(parentFolderId);
			const view = new ews.FolderView(100);

			const findResults = await this.service.FindFolders(parentFolder, view);

			const folders: IEwsFolder[] = [];
			for (const folder of findResults.Folders) {
				folders.push({
					FolderId: { Id: folder.Id.UniqueId },
					DisplayName: folder.DisplayName,
					TotalCount: folder.TotalCount,
					ChildFolderCount: folder.ChildFolderCount,
				});
			}

			return folders;
		} catch (error) {
			this.handleError(error, 'GetFolders');
		}
	}

	async updateFolder(folderId: string, updates: IDataObject): Promise<IEwsFolder> {
		try {
			const folderIdObj = new ews.FolderId(folderId);
			const folder = await ews.Folder.Bind(this.service, folderIdObj);

			if (updates.DisplayName) {
				folder.DisplayName = updates.DisplayName as string;
			}

			await folder.Update();

			return {
				FolderId: { Id: folder.Id.UniqueId },
				DisplayName: folder.DisplayName,
			};
		} catch (error) {
			this.handleError(error, 'UpdateFolder');
		}
	}

	async deleteFolder(folderId: string, hardDelete = false): Promise<void> {
		try {
			const folderIdObj = new ews.FolderId(folderId);
			const deleteMode = hardDelete
				? ews.DeleteMode.HardDelete
				: ews.DeleteMode.SoftDelete;

			const folder = await ews.Folder.Bind(this.service, folderIdObj);
			await folder.Delete(deleteMode);
		} catch (error) {
			this.handleError(error, 'DeleteFolder');
		}
	}

	// ===============================
	// CALENDAR/EVENT OPERATIONS (Simplified - Events are CalendarItems)
	// ===============================

	async createCalendar(calendarName: string): Promise<IEwsCalendar> {
		try {
			const calendarFolder = new ews.CalendarFolder(this.service);
			calendarFolder.DisplayName = calendarName;

			await calendarFolder.Save(ews.WellKnownFolderName.Calendar);

			return {
				FolderId: { Id: calendarFolder.Id.UniqueId },
				DisplayName: calendarFolder.DisplayName,
			};
		} catch (error) {
			this.handleError(error, 'CreateCalendar');
		}
	}

	async getCalendar(calendarId: string): Promise<IEwsCalendar> {
		return this.getFolder(calendarId);
	}

	async getCalendars(): Promise<IEwsCalendar[]> {
		try {
			// AWS WorkMail/Exchange has one default calendar per mailbox
			// Return the default calendar folder
			const calendar = await ews.Folder.Bind(
				this.service,
				ews.WellKnownFolderName.Calendar
			);

			return [{
				FolderId: { Id: calendar.Id.UniqueId },
				DisplayName: calendar.DisplayName,
				TotalCount: calendar.TotalCount,
				ChildFolderCount: calendar.ChildFolderCount,
			}];
		} catch (error) {
			this.handleError(error, 'GetCalendars');
		}
	}

	async updateCalendar(calendarId: string, updates: IDataObject): Promise<IEwsCalendar> {
		return this.updateFolder(calendarId, updates);
	}

	async deleteCalendar(calendarId: string): Promise<void> {
		return this.deleteFolder(calendarId, true);
	}

	async createEvent(calendarId: string, event: IDataObject): Promise<IEwsCalendarItem> {
		try {
			const appointment = new ews.Appointment(this.service);

			if (event.subject) {
				appointment.Subject = event.subject as string;
			}

			if (event.body) {
				const bodyType = event.bodyType === 'Text'
					? ews.BodyType.Text
					: ews.BodyType.HTML;
				appointment.Body = new ews.MessageBody(bodyType, event.body as string);
			}

			if (event.start) {
				appointment.Start = new ews.DateTime(new Date(event.start as string));
			}

			if (event.end) {
				appointment.End = new ews.DateTime(new Date(event.end as string));
			}

			if (event.location) {
				appointment.Location = event.location as string;
			}

			if (event.isAllDayEvent !== undefined) {
				appointment.IsAllDayEvent = event.isAllDayEvent as boolean;
			}

			// Add attendees
			if (event.requiredAttendees) {
				const attendees = Array.isArray(event.requiredAttendees)
					? event.requiredAttendees
					: [event.requiredAttendees];
				attendees.forEach((email: any) => {
					appointment.RequiredAttendees.Add(typeof email === 'string' ? email : email.EmailAddress);
				});
			}

			if (event.optionalAttendees) {
				const attendees = Array.isArray(event.optionalAttendees)
					? event.optionalAttendees
					: [event.optionalAttendees];
				attendees.forEach((email: any) => {
					appointment.OptionalAttendees.Add(typeof email === 'string' ? email : email.EmailAddress);
				});
			}

			if (calendarId === 'calendar') {
				await appointment.Save(ews.WellKnownFolderName.Calendar);
			} else {
				await appointment.Save(new ews.FolderId(calendarId));
			}

			return {
				ItemId: { Id: appointment.Id.UniqueId },
				Subject: appointment.Subject,
				Start: appointment.Start.toString(),
				End: appointment.End.toString(),
			};
		} catch (error) {
			this.handleError(error, 'CreateEvent');
		}
	}

	async getEvent(eventId: string): Promise<IEwsCalendarItem> {
		try {
			const itemId = new ews.ItemId(eventId);
			const appointment = await ews.Appointment.Bind(this.service, itemId);

			return {
				ItemId: { Id: appointment.Id.UniqueId },
				Subject: appointment.Subject,
				Start: appointment.Start.toString(),
				End: appointment.End.toString(),
				Location: appointment.Location,
				IsAllDayEvent: appointment.IsAllDayEvent,
			};
		} catch (error) {
			this.handleError(error, 'GetEvent');
		}
	}

	async getEvents(calendarId = 'calendar', options: IEwsSearchOptions = {}): Promise<IEwsCalendarItem[]> {
		try {
			const maxResults = options.maxResults || 50;
			const offset = options.offset || 0;

			const view = new ews.ItemView(maxResults, offset);
			let findResults;
			if (calendarId === 'calendar') {
				findResults = await this.service.FindItems(ews.WellKnownFolderName.Calendar, view);
			} else {
				findResults = await this.service.FindItems(new ews.FolderId(calendarId), view);
			}

			const events: IEwsCalendarItem[] = [];
			for (const item of findResults.Items) {
				if (item instanceof ews.Appointment) {
					events.push({
						ItemId: { Id: item.Id.UniqueId },
						Subject: item.Subject,
						Start: item.Start.toString(),
						End: item.End.toString(),
					});
				}
			}

			return events;
		} catch (error) {
			this.handleError(error, 'GetEvents');
		}
	}

	async updateEvent(eventId: string, updates: IDataObject): Promise<IEwsCalendarItem> {
		try {
			const itemId = new ews.ItemId(eventId);
			const appointment = await ews.Appointment.Bind(this.service, itemId);

			if (updates.Subject) {
				appointment.Subject = updates.Subject as string;
			}

			if (updates.Start) {
				appointment.Start = new ews.DateTime(new Date(updates.Start as string));
			}

			if (updates.End) {
				appointment.End = new ews.DateTime(new Date(updates.End as string));
			}

			if (updates.Location) {
				appointment.Location = updates.Location as string;
			}

			await appointment.Update(ews.ConflictResolutionMode.AlwaysOverwrite);

			return {
				ItemId: { Id: appointment.Id.UniqueId },
				Subject: appointment.Subject,
			};
		} catch (error) {
			this.handleError(error, 'UpdateEvent');
		}
	}

	async deleteEvent(eventId: string): Promise<void> {
		try {
			const itemId = new ews.ItemId(eventId);
			const appointment = await ews.Appointment.Bind(this.service, itemId);
			await appointment.Delete(ews.DeleteMode.MoveToDeletedItems);
		} catch (error) {
			this.handleError(error, 'DeleteEvent');
		}
	}

	// ===============================
	// CONTACT OPERATIONS
	// ===============================

	async createContact(contact: IDataObject): Promise<IEwsContact> {
		try {
			const ewsContact = new ews.Contact(this.service);

			if (contact.displayName) {
				ewsContact.DisplayName = contact.displayName as string;
			}

			if (contact.givenName) {
				ewsContact.GivenName = contact.givenName as string;
			}

			if (contact.surname) {
				ewsContact.Surname = contact.surname as string;
			}

			if (contact.emailAddress) {
				ewsContact.EmailAddresses._setItem(
					ews.EmailAddressKey.EmailAddress1,
					new ews.EmailAddress(contact.emailAddress as string)
				);
			}

			if (contact.companyName) {
				ewsContact.CompanyName = contact.companyName as string;
			}

			if (contact.jobTitle) {
				ewsContact.JobTitle = contact.jobTitle as string;
			}

			await ewsContact.Save(ews.WellKnownFolderName.Contacts);

			return {
				ItemId: { Id: ewsContact.Id.UniqueId },
				DisplayName: ewsContact.DisplayName,
			};
		} catch (error) {
			this.handleError(error, 'CreateContact');
		}
	}

	async getContact(contactId: string): Promise<IEwsContact> {
		try {
			const itemId = new ews.ItemId(contactId);
			const contact = await ews.Contact.Bind(this.service, itemId);

			return {
				ItemId: { Id: contact.Id.UniqueId },
				DisplayName: contact.DisplayName,
				GivenName: contact.GivenName,
				Surname: contact.Surname,
			};
		} catch (error) {
			this.handleError(error, 'GetContact');
		}
	}

	async getContacts(options: IEwsSearchOptions = {}): Promise<IEwsContact[]> {
		try {
			const maxResults = options.maxResults || 50;
			const offset = options.offset || 0;

			const view = new ews.ItemView(maxResults, offset);
			const findResults = await this.service.FindItems(ews.WellKnownFolderName.Contacts, view);

			const contacts: IEwsContact[] = [];
			for (const item of findResults.Items) {
				if (item instanceof ews.Contact) {
					contacts.push({
						ItemId: { Id: item.Id.UniqueId },
						DisplayName: item.DisplayName,
					});
				}
			}

			return contacts;
		} catch (error) {
			this.handleError(error, 'GetContacts');
		}
	}

	async updateContact(contactId: string, updates: IDataObject): Promise<IEwsContact> {
		try {
			const itemId = new ews.ItemId(contactId);
			const contact = await ews.Contact.Bind(this.service, itemId);

			if (updates.DisplayName) {
				contact.DisplayName = updates.DisplayName as string;
			}

			if (updates.GivenName) {
				contact.GivenName = updates.GivenName as string;
			}

			if (updates.Surname) {
				contact.Surname = updates.Surname as string;
			}

			await contact.Update(ews.ConflictResolutionMode.AlwaysOverwrite);

			return {
				ItemId: { Id: contact.Id.UniqueId },
				DisplayName: contact.DisplayName,
			};
		} catch (error) {
			this.handleError(error, 'UpdateContact');
		}
	}

	async deleteContact(contactId: string): Promise<void> {
		try {
			const itemId = new ews.ItemId(contactId);
			const contact = await ews.Contact.Bind(this.service, itemId);
			await contact.Delete(ews.DeleteMode.MoveToDeletedItems);
		} catch (error) {
			this.handleError(error, 'DeleteContact');
		}
	}

	// ===============================
	// ATTACHMENT OPERATIONS
	// ===============================

	async addAttachment(itemId: string, attachment: IDataObject): Promise<IEwsAttachment> {
		try {
			const item = await ews.Item.Bind(this.service, new ews.ItemId(itemId));

			// AddFileAttachment expects string (base64) or file path
			const fileAttachment = item.Attachments.AddFileAttachment(
				attachment.name as string,
				attachment.content as string
			);

			await item.Update(ews.ConflictResolutionMode.AlwaysOverwrite);

			return {
				AttachmentId: { Id: fileAttachment.Id },
				Name: fileAttachment.Name,
			};
		} catch (error) {
			this.handleError(error, 'AddAttachment');
		}
	}

	async getAttachment(attachmentId: string): Promise<IEwsAttachment> {
		try {
			// Note: Getting individual attachments requires loading from parent item
			// This is a simplified version
			return {
				AttachmentId: { Id: attachmentId },
				Name: 'Attachment',
			};
		} catch (error) {
			this.handleError(error, 'GetAttachment');
		}
	}

	async getAttachments(itemId: string): Promise<IEwsAttachment[]> {
		try {
			const item = await ews.Item.Bind(
				this.service,
				new ews.ItemId(itemId),
				new ews.PropertySet(ews.BasePropertySet.FirstClassProperties, [ews.ItemSchema.Attachments])
			);

			const attachments: IEwsAttachment[] = [];
			for (let i = 0; i < item.Attachments.Count; i++) {
				const attachment = item.Attachments._getItem(i);
				if (attachment instanceof ews.FileAttachment) {
					attachments.push({
						AttachmentId: { Id: attachment.Id },
						Name: attachment.Name,
						Size: attachment.Size,
					});
				}
			}

			return attachments;
		} catch (error) {
			this.handleError(error, 'GetAttachments');
		}
	}

	async downloadAttachment(_attachmentId: string): Promise<Buffer> {
		try {
			// Note: This requires binding to the parent item and loading the attachment
			// Simplified version - needs parent item ID in real implementation
			return Buffer.from('');
		} catch (error) {
			this.handleError(error, 'DownloadAttachment');
		}
	}

	// ===============================
	// HELPER METHODS
	// ===============================

	private getFolderId(folderId: string): any {
		// Map well-known folder names
		const wellKnownFolders: { [key: string]: any } = {
			'inbox': ews.WellKnownFolderName.Inbox,
			'sentitems': ews.WellKnownFolderName.SentItems,
			'deleteditems': ews.WellKnownFolderName.DeletedItems,
			'drafts': ews.WellKnownFolderName.Drafts,
			'junkemail': ews.WellKnownFolderName.JunkEmail,
			'msgfolderroot': ews.WellKnownFolderName.MsgFolderRoot,
			'calendar': ews.WellKnownFolderName.Calendar,
			'contacts': ews.WellKnownFolderName.Contacts,
		};

		const lowerFolderId = folderId.toLowerCase();
		if (wellKnownFolders[lowerFolderId]) {
			return wellKnownFolders[lowerFolderId];
		}

		return new ews.FolderId(folderId);
	}

	private convertMessageToJson(message: ews.EmailMessage): IEwsMessage {
		// Extract Body safely - try multiple approaches
		let bodyData;
		try {
			let bodyValue = '';
			let bodyType = 'Text';
			
			if (message.Body) {
				// Versuche verschiedene Wege, den Body-Text zu extrahieren
				if (typeof message.Body.Text === 'string' && message.Body.Text.length > 0) {
					bodyValue = message.Body.Text;
				} else if (typeof (message.Body as any).text === 'string' && (message.Body as any).text.length > 0) {
					bodyValue = (message.Body as any).text;
				} else if (typeof (message.Body as any).Content === 'string' && (message.Body as any).Content.length > 0) {
					bodyValue = (message.Body as any).Content;
				} else if (typeof (message.Body as any).content === 'string' && (message.Body as any).content.length > 0) {
					bodyValue = (message.Body as any).content;
				} else if (typeof message.Body.toString === 'function') {
					const strValue = message.Body.toString();
					// Nur verwenden wenn es kein "[object Object]" ist
					if (strValue && !strValue.includes('[object') && strValue.length > 0) {
						bodyValue = strValue;
					}
				}
				
				// Body-Type extrahieren
				if (message.Body.BodyType !== undefined && message.Body.BodyType !== null) {
					if (typeof message.Body.BodyType === 'number') {
						bodyType = message.Body.BodyType === 0 ? 'HTML' : 'Text';
					} else {
						bodyType = String(message.Body.BodyType);
					}
				}
			}
			
			bodyData = {
				BodyType: bodyType,
				Value: bodyValue,
			};
		} catch {
			// Body not loaded
			bodyData = {
				BodyType: 'Text',
				Value: '',
			};
		}

		return {
			ItemId: { Id: message.Id?.UniqueId || '' },
			Subject: message.Subject,
			Body: bodyData,
			From: message.From ? {
				Name: message.From.Name,
				EmailAddress: message.From.Address,
			} : undefined,
			IsRead: message.IsRead,
			DateTimeReceived: message.DateTimeReceived?.toString(),
			DateTimeSent: message.DateTimeSent?.toString(),
			HasAttachments: message.HasAttachments,
		};
	}

	private mapImportance(importance: string): ews.Importance {
		switch (importance.toLowerCase()) {
			case 'low':
				return ews.Importance.Low;
			case 'high':
				return ews.Importance.High;
			default:
				return ews.Importance.Normal;
		}
	}

	private mapSensitivity(sensitivity: string): ews.Sensitivity {
		switch (sensitivity.toLowerCase()) {
			case 'personal':
				return ews.Sensitivity.Personal;
			case 'private':
				return ews.Sensitivity.Private;
			case 'confidential':
				return ews.Sensitivity.Confidential;
			default:
				return ews.Sensitivity.Normal;
		}
	}
}
