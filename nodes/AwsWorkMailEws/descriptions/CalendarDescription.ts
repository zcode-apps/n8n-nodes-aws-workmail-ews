import { INodeProperties } from 'n8n-workflow';

export const calendarOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['calendar'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a calendar',
				action: 'Create a calendar',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a calendar',
				action: 'Delete a calendar',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a calendar',
				action: 'Get a calendar',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many calendars',
				action: 'Get many calendars',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a calendar',
				action: 'Update a calendar',
			},
		],
		default: 'getAll',
	},
];

export const calendarFields: INodeProperties[] = [
	// ----------------------------------
	//         calendar:create
	// ----------------------------------
	{
		displayName: 'Calendar Name',
		name: 'calendarName',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['calendar'],
				operation: ['create'],
			},
		},
		description: 'Name of the calendar to create',
	},

	// ----------------------------------
	//         calendar:get/delete/update
	// ----------------------------------
	{
		displayName: 'Calendar ID',
		name: 'calendarId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['calendar'],
				operation: ['get', 'delete', 'update'],
			},
		},
		description: 'ID of the calendar',
	},

	// ----------------------------------
	//         calendar:update
	// ----------------------------------
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['calendar'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Display Name',
				name: 'DisplayName',
				type: 'string',
				default: '',
				description: 'New name for the calendar',
			},
		],
	},
];
