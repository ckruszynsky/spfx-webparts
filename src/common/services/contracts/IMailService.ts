import { MailMessage } from '../../models/mail';

export interface IMailService {
    getMailForCurrentUser:()=> Promise<Array<MailMessage>>;
}