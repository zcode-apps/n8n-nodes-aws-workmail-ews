import { INodeProperties } from 'n8n-workflow';

export const eventOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['event'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create an event',
				action: 'Create an event',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete an event',
				action: 'Delete an event',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get an event',
				action: 'Get an event',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many events',
				action: 'Get many events',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update an event',
				action: 'Update an event',
			},
		],
		default: 'create',
	},
];

export const eventFields: INodeProperties[] = [
	// ----------------------------------
	//         event:create
	// ----------------------------------
	{
		displayName: 'Calendar ID',
		name: 'calendarId',
		type: 'string',
		default: 'calendar',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description: 'ID of the calendar (use "calendar" for default calendar)',
	},
	{
		displayName: 'Subject',
		name: 'subject',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description: 'Subject/title of the event',
	},
	{
		displayName: 'Start Time',
		name: 'start',
		type: 'dateTime',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description: 'Start time of the event',
	},
	{
		displayName: 'End Time',
		name: 'end',
		type: 'dateTime',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description: 'End time of the event',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		options: [
			{
				displayName: 'Body',
				name: 'body',
				type: 'string',
				typeOptions: {
					rows: 4,
				},
				default: '',
				description: 'Body/description of the event',
			},
			{
				displayName: 'Body Type',
				name: 'bodyType',
				type: 'options',
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
				description: 'Format of the event body',
			},
			{
				displayName: 'Location',
				name: 'location',
				type: 'string',
				default: '',
				description: 'Location of the event',
			},
			{
				displayName: 'Is All Day Event',
				name: 'isAllDayEvent',
				type: 'boolean',
				default: false,
				description: 'Whether the event is an all-day event',
			},
			{
				displayName: 'Required Attendees',
				name: 'requiredAttendees',
				type: 'string',
				default: '',
				description: 'Email addresses of required attendees. Multiple can be added separated by comma.',
			},
			{
				displayName: 'Optional Attendees',
				name: 'optionalAttendees',
				type: 'string',
				default: '',
				description: 'Email addresses of optional attendees. Multiple can be added separated by comma.',
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
				description: 'Importance of the event',
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
				description: 'Sensitivity of the event',
			},
			{
				displayName: 'Categories',
				name: 'categories',
				type: 'string',
				default: '',
				description: 'Categories for the event. Multiple can be added separated by comma.',
			},
		],
	},

	// ----------------------------------
	//         event:get/delete
	// ----------------------------------
	{
		displayName: 'Event ID',
		name: 'eventId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['get', 'delete', 'update'],
			},
		},
		description: 'ID of the event',
	},

	// ----------------------------------
	//         event:getAll
	// ----------------------------------
	{
		displayName: 'Calendar ID',
		name: 'calendarId',
		type: 'string',
		default: 'calendar',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['getAll'],
			},
		},
		description: 'ID of the calendar to get events from',
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['event'],
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
				resource: ['event'],
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

	// ----------------------------------
	//         event:update
	// ----------------------------------
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Subject',
				name: 'Subject',
				type: 'string',
				default: '',
				description: 'New subject/title for the event',
			},
			{
				displayName: 'Start Time',
				name: 'Start',
				type: 'dateTime',
				default: '',
				description: 'New start time for the event',
			},
			{
				displayName: 'End Time',
				name: 'End',
				type: 'dateTime',
				default: '',
				description: 'New end time for the event',
			},
			{
				displayName: 'Location',
				name: 'Location',
				type: 'string',
				default: '',
				description: 'New location for the event',
			},
			{
				displayName: 'Body',
				name: 'Body',
				type: 'string',
				typeOptions: {
					rows: 4,
				},
				default: '',
				description: 'New body/description for the event',
			},
		],
	},
];
