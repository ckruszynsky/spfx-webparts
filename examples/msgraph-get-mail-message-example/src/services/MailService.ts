import {ServiceKey, ServiceScope} from '@microsoft/sp-core-library';
import {MSGraphClientFactory, MSGraphClient} from '@microsoft/sp-http';
import { MailMessage } from '../models';

export class MailService {
    private ENDPOINT_URL:string = "https://graph.microsoft.com/v1.0/me/messages";
    

    constructor(private graphClientFactory:MSGraphClientFactory){}

    public async getMail():Promise<Array<MailMessage>>{
        var client = await this.graphClientFactory.getClient();
        var clientMailResponse = await client.api(this.ENDPOINT_URL).get();        
        var messages = new Array<MailMessage>();
        clientMailResponse.value.forEach(value => {
            let message:MailMessage = {
                id: value.id,
                createdDate:value.createdDate,
                lastModifiedDate:value.lastModifiedDate,
                receivedDateTime: value.receivedDateTime,
                sentDateTime: value.sentDateTime,
                hasAttachments:value.hasAttachments,
                subject:value.subject,
                bodyPreview:value.bodyPreview,
                importance:value.importance,
                isRead: value.isRead,
                isDraft:value.isDraft,
                body:value.body,
                sender: {
                    emailAddress: {
                        name: value.sender.emailAddress.name,
                        address: value.sender.emailAddress.address
                    }
                },
                from: {
                    emailAddress: {
                        name: value.sender.emailAddress.name,
                        address: value.sender.emailAddress.address
                    }
                }
            };

            messages.push(message);
        });

        messages.sort((m1,m2)=> 
            m1.receivedDateTime > m2.receivedDateTime ? -1: 1);
        
        console.log(messages);
        return messages;
    }
}