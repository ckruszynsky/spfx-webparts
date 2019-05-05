import { MailRecipient } from "./MailRecipient";

export interface MailMessage {
    id:string;
    createdDate:Date;
    lastModifiedDate:Date;
    receivedDateTime?:Date;
    sentDateTime?:Date;
    hasAttachments:boolean;
    subject:string;
    bodyPreview:string;
    importance:"string";
    isRead:boolean;
    isDraft:boolean;
    body:string;
    sender:MailRecipient,
    from:MailRecipient  
}