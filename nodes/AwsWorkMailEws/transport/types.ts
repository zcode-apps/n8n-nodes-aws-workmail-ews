import { IDataObject } from 'n8n-workflow';

export interface IEwsCredentials {
	ewsUrl: string;
	username: string;
	password: string;
}

export interface IEwsMessage extends IDataObject {
	ItemId?: {
		Id: string;
		ChangeKey?: string;
	};
	Subject?: string;
	Body?: {
		BodyType: string;
		Value: string;
	};
	From?: IEwsEmailAddress;
	ToRecipients?: IEwsEmailAddress[];
	CcRecipients?: IEwsEmailAddress[];
	BccRecipients?: IEwsEmailAddress[];
	IsRead?: boolean;
	DateTimeReceived?: string;
	DateTimeSent?: string;
	HasAttachments?: boolean;
	Attachments?: IEwsAttachment[];
	Importance?: string;
	Sensitivity?: string;
	Categories?: string[];
	ParentFolderId?: {
		Id: string;
		ChangeKey?: string;
	};
}

export interface IEwsEmailAddress extends IDataObject {
	Name?: string;
	EmailAddress: string;
	RoutingType?: string;
	MailboxType?: string;
}

export interface IEwsFolder extends IDataObject {
	FolderId?: {
		Id: string;
		ChangeKey?: string;
	};
	DisplayName?: string;
	TotalCount?: number;
	ChildFolderCount?: number;
	UnreadCount?: number;
	FolderClass?: string;
	ParentFolderId?: {
		Id: string;
		ChangeKey?: string;
	};
}

export interface IEwsCalendar extends IDataObject {
	FolderId?: {
		Id: string;
		ChangeKey?: string;
	};
	DisplayName?: string;
	TotalCount?: number;
	ChildFolderCount?: number;
}

export interface IEwsCalendarItem extends IDataObject {
	ItemId?: {
		Id: string;
		ChangeKey?: string;
	};
	Subject?: string;
	Body?: {
		BodyType: string;
		Value: string;
	};
	Start?: string;
	End?: string;
	StartTimeZone?: string;
	EndTimeZone?: string;
	Location?: string;
	Organizer?: IEwsEmailAddress;
	RequiredAttendees?: IEwsAttendee[];
	OptionalAttendees?: IEwsAttendee[];
	IsAllDayEvent?: boolean;
	LegacyFreeBusyStatus?: string;
	Importance?: string;
	Sensitivity?: string;
	IsRecurring?: boolean;
	IsCancelled?: boolean;
	Categories?: string[];
}

export interface IEwsAttendee extends IDataObject {
	Mailbox: IEwsEmailAddress;
	ResponseType?: string;
}

export interface IEwsContact extends IDataObject {
	ItemId?: {
		Id: string;
		ChangeKey?: string;
	};
	DisplayName?: string;
	GivenName?: string;
	Surname?: string;
	EmailAddresses?: {
		Entry: Array<{
			Key: string;
			Value: string;
		}>;
	};
	PhoneNumbers?: {
		Entry: Array<{
			Key: string;
			Value: string;
		}>;
	};
	CompanyName?: string;
	JobTitle?: string;
	PhysicalAddresses?: IDataObject;
	Categories?: string[];
}

export interface IEwsAttachment extends IDataObject {
	AttachmentId?: {
		Id: string;
		RootItemId?: string;
		RootItemChangeKey?: string;
	};
	Name?: string;
	ContentType?: string;
	ContentId?: string;
	ContentLocation?: string;
	Size?: number;
	IsInline?: boolean;
	IsContactPhoto?: boolean;
	Content?: string; // Base64 encoded
}

export interface IEwsSearchOptions {
	maxResults?: number;
	offset?: number;
	sort?: string;
	filters?: IDataObject;
}
