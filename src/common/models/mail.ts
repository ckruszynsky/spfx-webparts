export type MailAddress = {
    name:string;
    address:string;
}

export type MailRecipient = {
    emailAddress: MailAddress;
}


export type MailMessage ={
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