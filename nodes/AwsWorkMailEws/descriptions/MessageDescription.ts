import { INodeProperties } from 'n8n-workflow';

export const messageOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['message'],
			},
		},
		options: [
			{
				name: 'Create Reply Draft',
				value: 'createReplyDraft',
				description: 'Create a reply draft that can be reviewed before sending',
				action: 'Create a reply draft',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a message',
				action: 'Delete a message',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a message',
				action: 'Get a message',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many messages',
				action: 'Get many messages',
			},
			{
				name: 'Move',
				value: 'move',
				description: 'Move a message to a folder',
				action: 'Move a message',
			},
			{
				name: 'Reply',
				value: 'reply',
				description: 'Reply to a message',
				action: 'Reply to a message',
			},
			{
				name: 'Send',
				value: 'send',
				description: 'Send a message',
				action: 'Send a message',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a message',
				action: 'Update a message',
			},
		],
		default: 'send',
	},
];

export const messageFields: INodeProperties[] = [
	// ----------------------------------
	//         message:send
	// ----------------------------------
	{
		displayName: 'To Recipients',
		name: 'toRecipients',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['send'],
			},
		},
		description: 'Email addresses of recipients. Multiple can be added separated by comma.',
	},
	{
		displayName: 'Subject',
		name: 'subject',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['send'],
			},
		},
		description: 'Subject of the message',
	},
	{
		displayName: 'Body',
		name: 'body',
		type: 'string',
		typeOptions: {
			rows: 5,
		},
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['send'],
			},
		},
		description: 'Body of the message',
	},
	{
		displayName: 'Body Type',
		name: 'bodyType',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['send'],
			},
		},
		options: [
			{
				name: 'HTML',
				value: 'HTML',
			},
			{
				name: 'Text',
				value: 'Text',
			},
		],
		default: 'HTML',
		description: 'The format of the message body',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['send'],
			},
		},
		options: [
			{
				displayName: 'BCC Recipients',
				name: 'bccRecipients',
				type: 'string',
				default: '',
				description: 'Email addresses of BCC recipients. Multiple can be added separated by comma.',
			},
			{
				displayName: 'CC Recipients',
				name: 'ccRecipients',
				type: 'string',
				default: '',
				description: 'Email addresses of CC recipients. Multiple can be added separated by comma.',
			},
			{
				displayName: 'Categories',
				name: 'categories',
				type: 'string',
				default: '',
				description: 'Categories associated with the message. Multiple can be added separated by comma.',
			},
			{
				displayName: 'Importance',
				name: 'importance',
				type: 'options',
				options: [
					{
						name: 'Low',
						value: 'Low',
					},
					{
						name: 'Normal',
						value: 'Normal',
					},
					{
						name: 'High',
						value: 'High',
					},
				],
				default: 'Normal',
				description: 'Importance of the message',
			},
			{
				displayName: 'Sensitivity',
				name: 'sensitivity',
				type: 'options',
				options: [
					{
						name: 'Normal',
						value: 'Normal',
					},
					{
						name: 'Personal',
						value: 'Personal',
					},
					{
						name: 'Private',
						value: 'Private',
					},
					{
						name: 'Confidential',
						value: 'Confidential',
					},
				],
				default: 'Normal',
				description: 'Sensitivity of the message',
			},
		],
	},

	// ----------------------------------
	//         message:get
	// ----------------------------------
	{
		displayName: 'Message ID',
		name: 'messageId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['get', 'delete', 'update', 'move', 'reply', 'createReplyDraft'],
			},
		},
		description: 'ID of the message',
	},

	// ----------------------------------
	//         message:getAll
	// ----------------------------------
	{
		displayName: 'Folder',
		name: 'folderId',
		type: 'string',
		default: 'inbox',
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['getAll'],
			},
		},
		description: 'Folder ID or distinguished folder name (inbox, sentitems, deleteditems, drafts)',
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['getAll'],
			},
		},
		default: false,
		description: 'Whether to return all results or only up to a given limit',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['getAll'],
				returnAll: [false],
			},
		},
		typeOptions: {
			minValue: 1,
			maxValue: 500,
		},
		default: 50,
		description: 'Max number of results to return',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				displayName: 'Sort',
				name: 'sort',
				type: 'options',
				options: [
					{
						name: 'Newest First',
						value: 'desc',
					},
					{
						name: 'Oldest First',
						value: 'asc',
					},
				],
				default: 'desc',
				description: 'Sort order for messages',
			},
		],
	},

	// ----------------------------------
	//         message:update
	// ----------------------------------
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Is Read',
				name: 'IsRead',
				type: 'boolean',
				default: false,
				description: 'Whether the message has been read',
			},
			{
				displayName: 'Subject',
				name: 'Subject',
				type: 'string',
				default: '',
				description: 'Subject of the message',
			},
			{
				displayName: 'Importance',
				name: 'Importance',
				type: 'options',
				options: [
					{
						name: 'Low',
						value: 'Low',
					},
					{
						name: 'Normal',
						value: 'Normal',
					},
					{
						name: 'High',
						value: 'High',
					},
				],
				default: 'Normal',
				description: 'Importance of the message',
			},
			{
				displayName: 'Sensitivity',
				name: 'Sensitivity',
				type: 'options',
				options: [
					{
						name: 'Normal',
						value: 'Normal',
					},
					{
						name: 'Personal',
						value: 'Personal',
					},
					{
						name: 'Private',
						value: 'Private',
					},
					{
						name: 'Confidential',
						value: 'Confidential',
					},
				],
				default: 'Normal',
				description: 'Sensitivity of the message',
			},
		],
	},

	// ----------------------------------
	//         message:move
	// ----------------------------------
	{
		displayName: 'Target Folder ID',
		name: 'targetFolderId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['move'],
			},
		},
		description: 'ID of the folder to move the message to',
	},

	// ----------------------------------
	//         message:reply
	// ----------------------------------
	{
		displayName: 'Reply Body',
		name: 'replyBody',
		type: 'string',
		typeOptions: {
			rows: 5,
		},
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['reply'],
			},
		},
		description: 'Body of the reply',
	},
	{
		displayName: 'Reply All',
		name: 'replyAll',
		type: 'boolean',
		default: false,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['reply'],
			},
		},
		description: 'Whether to reply to all recipients',
	},

	// ----------------------------------
	//         message:createReplyDraft
	// ----------------------------------
	{
		displayName: 'Reply Body',
		name: 'draftReplyBody',
		type: 'string',
		typeOptions: {
			rows: 5,
		},
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['createReplyDraft'],
			},
		},
		description: 'Body of the reply draft. This will be saved as a draft in the Drafts folder.',
	},
	{
		displayName: 'Body Type',
		name: 'draftBodyType',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['createReplyDraft'],
			},
		},
		options: [
			{
				name: 'HTML',
				value: 'HTML',
			},
			{
				name: 'Text',
				value: 'Text',
			},
		],
		default: 'HTML',
		description: 'The format of the reply body',
	},
	{
		displayName: 'Reply All',
		name: 'draftReplyAll',
		type: 'boolean',
		default: false,
		displayOptions: {
			show: {
				resource: ['message'],
				operation: ['createReplyDraft'],
			},
		},
		description: 'Whether to include all original recipients in the reply draft',
	},
];
