//https://graph.microsoft.com/v1.0/me/messages?$select=subject,bodyPreview,sender,webLink&$top=5
export interface EmailAddress {
    name: string;
    address: string;
}

export interface Sender {
    emailAddress: EmailAddress;
}

export interface Mail {
    subject: string;
    bodyPreview: string;
    webLink: string;
    sender: Sender;
}

export interface MailsCardData {
    title: string;
    mails: Mail[];
}