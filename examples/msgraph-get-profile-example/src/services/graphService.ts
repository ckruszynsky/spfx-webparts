import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import * as strings from 'MsGraphUserProfileWebPartStrings';

export class GraphService {
    constructor(private graphClientFactory:MSGraphClientFactory){}

    public async getUserProfile():Promise<MicrosoftGraph.User>{
        const userProfileClient = await this.graphClientFactory.getClient();
        const userProfileClientResponse = await userProfileClient.api(strings.GraphServiceURI).get();                
        return userProfileClientResponse;        
    }

    public getProfilePhoto(userName:string ):string {        
        return `/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId=${userName}`;
    }
}