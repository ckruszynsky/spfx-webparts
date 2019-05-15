import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory } from '@microsoft/sp-http';

import { UserProfile } from '../models';
import { IUserProfileService } from './contracts/IUserProfileService';


export default class UserProfileService implements IUserProfileService {
  public static readonly serviceKey: ServiceKey<IUserProfileService> = ServiceKey.create<IUserProfileService>(
    "spfx-webparts:IUsersService",
    UserProfileService
  );

  private _msGraphClientFactory: MSGraphClientFactory;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
    });
  }

  public async getUserPhoto(username: string): Promise<string> {
    return `/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId=${username}`;
  }

  public async getCurrentUser(): Promise<UserProfile> {
    try {
      const client = await this._msGraphClientFactory.getClient();
      const graphUser = await client.api("me").get();
      let user: UserProfile = {
        displayName: graphUser.displayName,
        jobTitle: graphUser.jobTitle,
        emailAddress: graphUser.mail,
        businessPhone: graphUser.businessPhones.length > 0 ? graphUser.businessPhones[0] : ""
      };
      return user;
    } catch (e) {
      console.error(e);
    }
  }
}
