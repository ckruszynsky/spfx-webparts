import {ServiceKey, ServiceScope} from '@microsoft/sp-core-library';
import {MSGraphClientFactory, MSGraphClient} from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export class EventService {
    private ENDPOINT_URL:string = "https://graph.microsoft.com/v1.0/me/events";
    
    constructor(private graphClientFactory:MSGraphClientFactory){}

    public async getEvents():Promise<MicrosoftGraph.Event[]>{
        var client = await this.graphClientFactory.getClient();
        var clientResponse = await client.api(this.ENDPOINT_URL).get();        
        return clientResponse.value;
    }
}