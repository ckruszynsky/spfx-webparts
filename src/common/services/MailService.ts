import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory } from '@microsoft/sp-http';

import { MailMessage } from '../models/mail';
import { IMailService } from './contracts/IMailService';

export default class MailService implements IMailService {
  public static readonly serviceKey: ServiceKey<IMailService> = ServiceKey.create<IMailService>(
    "spfx-webparts:IMailService",
    MailService
  );
  private _msGraphClientFactory: MSGraphClientFactory;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
    });
  }

  public async getMailForCurrentUser():Promise<Array<MailMessage>>{
      try{
        const client = await this._msGraphClientFactory.getClient();
        const response = await client.api('/me/messages').get();  
        
        let messages = response.value
             .map(msg => mapResponseToMessage(msg) as MailMessage)
             .sort(sortMessagesByReceivedDate);                
        return messages;
      }
      catch(e){
          console.error(e);
      }

  }
}
function sortMessagesByReceivedDate(m1:MailMessage,m2:MailMessage): number {
    return  m1.receivedDateTime > m2.receivedDateTime ? -1 : 1;
}

function mapResponseToMessage(msg: any): MailMessage {
    return {
        id: msg.id,
        createdDate: msg.createdDate,
        lastModifiedDate: msg.lastModifiedDate,
        receivedDateTime: msg.receivedDateTime,
        sentDateTime: msg.sentDateTime,
        hasAttachments: msg.hasAttachments,
        subject: msg.subject,
        bodyPreview: msg.bodyPreview,
        importance: msg.importance,
        isRead: msg.isRead,
        isDraft: msg.isDraft,
        body: msg.body,
        sender: {
            emailAddress: {
                name: msg.sender.emailAddress.name,
                address: msg.sender.emailAddress.address
            }
        },
        from: {
            emailAddress: {
                name: msg.sender.emailAddress.name,
                address: msg.sender.emailAddress.address
            }
        }
    };
}

