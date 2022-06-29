import { Message } from "@microsoft/microsoft-graph-types";

export interface MailsCardData {
    title: string;
    mails: Message[];
}