import { INodeProperties } from 'n8n-workflow';

export const folderOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a folder',
				action: 'Create a folder',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a folder',
				action: 'Delete a folder',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a folder',
				action: 'Get a folder',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many folders',
				action: 'Get many folders',
			},
			{
				name: 'Get Many Folder Messages',
				value: 'getFolderMessages',
				description: 'Get messages from a specific folder',
				action: 'Get many folder messages',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a folder',
				action: 'Update a folder',
			},
		],
		default: 'getAll',
	},
];

export const folderFields: INodeProperties[] = [
	// ----------------------------------
	//         folder:create
	// ----------------------------------
	{
		displayName: 'Folder Name',
		name: 'folderName',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create'],
			},
		},
		description: 'Name of the folder to create',
	},
	{
		displayName: 'Parent Folder ID',
		name: 'parentFolderId',
		type: 'string',
		default: 'msgfolderroot',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create'],
			},
		},
		description: 'ID of the parent folder (use "msgfolderroot" for root level, "inbox" for inbox subfolder)',
	},

	// ----------------------------------
	//         folder:get/delete/update
	// ----------------------------------
	{
		displayName: 'Folder ID',
		name: 'folderId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['get', 'delete', 'update'],
			},
		},
		description: 'ID of the folder',
	},

	// ----------------------------------
	//         folder:getAll
	// ----------------------------------
	{
		displayName: 'Parent Folder ID',
		name: 'parentFolderId',
		type: 'string',
		default: 'msgfolderroot',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['getAll'],
			},
		},
		description: 'ID of the parent folder to list subfolders from',
	},

	// ----------------------------------
	//         folder:update
	// ----------------------------------
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Display Name',
				name: 'DisplayName',
				type: 'string',
				default: '',
				description: 'New name for the folder',
			},
		],
	},

	// ----------------------------------
	//         folder:getFolderMessages
	// ----------------------------------
	{
		displayName: 'Folder ID',
		name: 'folderId',
		type: 'string',
		default: 'inbox',
		required: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['getFolderMessages'],
			},
		},
		description: 'ID of the folder to get messages from',
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['getFolderMessages'],
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
				resource: ['folder'],
				operation: ['getFolderMessages'],
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
];
