import {ServiceKey, ServiceScope} from '@microsoft/sp-core-library';
import {MSGraphClientFactory, MSGraphClient} from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export class GroupService {
    private ENDPOINT_URL:string = "https://graph.microsoft.com/v1.0/me/memberOf";
    
    constructor(private graphClientFactory:MSGraphClientFactory){}

    public async getGroups():Promise<MicrosoftGraph.Group[]>{
        var client = await this.graphClientFactory.getClient();
        var clientResponse = await client.api(this.ENDPOINT_URL).get();        
        var values:Array<MicrosoftGraph.DirectoryObject> = clientResponse.value;
        var groups = values.filter((x) => x["@odata.type"] == "#microsoft.graph.group");
        return groups;
    }
}