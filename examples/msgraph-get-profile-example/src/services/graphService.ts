import {ServiceKey, ServiceScope} from '@microsoft/sp-core-library';
import {MSGraphClientFactory, MSGraphClient} from '@microsoft/sp-http';

export class GraphService {
    private ENDPOINT_URL:string = "https://graph.microsoft.com/v1.0/me";
    

    constructor(private graphClientFactory:MSGraphClientFactory){}

    public async getUserProfile():Promise<any>{
        var client = await this.graphClientFactory.getClient();
        var clientProfileResponse = await client.api(this.ENDPOINT_URL).get();        
        
        return {
            profile: clientProfileResponse,
            photo: `/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId=${clientProfileResponse.userPrincipalName} `
        };
    }
}