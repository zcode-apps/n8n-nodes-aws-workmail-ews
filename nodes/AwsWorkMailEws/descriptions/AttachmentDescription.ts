import { INodeProperties } from 'n8n-workflow';

export const attachmentOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['attachment'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Add an attachment to a message',
				action: 'Add an attachment',
			},
			{
				name: 'Download',
				value: 'download',
				description: 'Download an attachment',
				action: 'Download an attachment',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get an attachment',
				action: 'Get an attachment',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get all attachments of a message',
				action: 'Get many attachments',
			},
		],
		default: 'add',
	},
];

export const attachmentFields: INodeProperties[] = [
	// ----------------------------------
	//         attachment:add
	// ----------------------------------
	{
		displayName: 'Message ID',
		name: 'messageId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['attachment'],
				operation: ['add'],
			},
		},
		description: 'ID of the message to add attachment to',
	},
	{
		displayName: 'Binary Property',
		name: 'binaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['attachment'],
				operation: ['add'],
			},
		},
		description: 'Name of the binary property containing file to attach',
	},

	// ----------------------------------
	//         attachment:get/download
	// ----------------------------------
	{
		displayName: 'Attachment ID',
		name: 'attachmentId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['attachment'],
				operation: ['get', 'download'],
			},
		},
		description: 'ID of the attachment',
	},
	{
		displayName: 'Binary Property',
		name: 'binaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['attachment'],
				operation: ['download'],
			},
		},
		description: 'Name of the binary property to store the downloaded file',
	},

	// ----------------------------------
	//         attachment:getAll
	// ----------------------------------
	{
		displayName: 'Message ID',
		name: 'messageId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['attachment'],
				operation: ['getAll'],
			},
		},
		description: 'ID of the message to get attachments from',
	},
];
