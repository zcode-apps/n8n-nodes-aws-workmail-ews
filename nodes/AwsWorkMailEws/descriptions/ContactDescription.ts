import { INodeProperties } from 'n8n-workflow';

export const contactOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['contact'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a contact',
				action: 'Create a contact',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a contact',
				action: 'Delete a contact',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a contact',
				action: 'Get a contact',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many contacts',
				action: 'Get many contacts',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a contact',
				action: 'Update a contact',
			},
		],
		default: 'create',
	},
];

export const contactFields: INodeProperties[] = [
	// ----------------------------------
	//         contact:create
	// ----------------------------------
	{
		displayName: 'Display Name',
		name: 'displayName',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['contact'],
				operation: ['create'],
			},
		},
		description: 'Display name of the contact',
	},
	{
		displayName: 'Email Address',
		name: 'emailAddress',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['contact'],
				operation: ['create'],
			},
		},
		description: 'Email address of the contact',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['contact'],
				operation: ['create'],
			},
		},
		options: [
			{
				displayName: 'Given Name',
				name: 'givenName',
				type: 'string',
				default: '',
				description: 'First name of the contact',
			},
			{
				displayName: 'Surname',
				name: 'surname',
				type: 'string',
				default: '',
				description: 'Last name of the contact',
			},
			{
				displayName: 'Company Name',
				name: 'companyName',
				type: 'string',
				default: '',
				description: 'Company of the contact',
			},
			{
				displayName: 'Job Title',
				name: 'jobTitle',
				type: 'string',
				default: '',
				description: 'Job title of the contact',
			},
			{
				displayName: 'Phone Number',
				name: 'phoneNumber',
				type: 'string',
				default: '',
				description: 'Phone number of the contact',
			},
			{
				displayName: 'Categories',
				name: 'categories',
				type: 'string',
				default: '',
				description: 'Categories for the contact. Multiple can be added separated by comma.',
			},
		],
	},

	// ----------------------------------
	//         contact:get/delete
	// ----------------------------------
	{
		displayName: 'Contact ID',
		name: 'contactId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['contact'],
				operation: ['get', 'delete', 'update'],
			},
		},
		description: 'ID of the contact',
	},

	// ----------------------------------
	//         contact:getAll
	// ----------------------------------
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['contact'],
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
				resource: ['contact'],
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
	//         contact:update
	// ----------------------------------
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['contact'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Display Name',
				name: 'DisplayName',
				type: 'string',
				default: '',
				description: 'New display name for the contact',
			},
			{
				displayName: 'Given Name',
				name: 'GivenName',
				type: 'string',
				default: '',
				description: 'New first name for the contact',
			},
			{
				displayName: 'Surname',
				name: 'Surname',
				type: 'string',
				default: '',
				description: 'New last name for the contact',
			},
		],
	},
];
