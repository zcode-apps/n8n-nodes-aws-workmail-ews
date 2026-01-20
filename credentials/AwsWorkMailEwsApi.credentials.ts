import {
	IAuthenticateGeneric,
	ICredentialTestRequest,
	ICredentialType,
	INodeProperties,
} from 'n8n-workflow';

export class AwsWorkMailEwsApi implements ICredentialType {
	name = 'awsWorkMailEwsApi';
	displayName = 'AWS WorkMail EWS API';
	documentationUrl = 'https://docs.aws.amazon.com/workmail/';
	properties: INodeProperties[] = [
		{
			displayName: 'EWS Endpoint URL',
			name: 'ewsUrl',
			type: 'string',
			default: '',
			placeholder: 'https://mobile.mail.eu-west-1.awsapps.com/EWS/Exchange.asmx',
			description: 'The Exchange Web Services endpoint URL for your AWS WorkMail organization',
			required: true,
		},
		{
			displayName: 'Email Address',
			name: 'username',
			type: 'string',
			default: '',
			placeholder: 'user@example.awsapps.com',
			description: 'Your AWS WorkMail email address',
			required: true,
		},
		{
			displayName: 'Password',
			name: 'password',
			type: 'string',
			typeOptions: {
				password: true,
			},
			default: '',
			description: 'Your AWS WorkMail password',
			required: true,
		},
	];

	authenticate: IAuthenticateGeneric = {
		type: 'generic',
		properties: {},
	};

	test: ICredentialTestRequest = {
		request: {
			baseURL: '={{$credentials.ewsUrl}}',
			url: '',
			method: 'POST',
			headers: {
				'Content-Type': 'text/xml; charset=utf-8',
				'SOAPAction': 'http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder',
			},
			body: `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2010_SP2" />
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="inbox" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>`,
			auth: {
				username: '={{$credentials.username}}',
				password: '={{$credentials.password}}',
			},
		},
	};
}
