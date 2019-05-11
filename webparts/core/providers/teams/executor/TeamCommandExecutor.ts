import { MSGraphClientFactory } from '@microsoft/sp-http';

import { Team } from '../models/Team';

export interface ITeamCommandExecutor {
    GetJoinedTeams: () => Promise<Team[]>
}

export class TeamCommandExecutor implements ITeamCommandExecutor {
    constructor(private graphClientFactory:MSGraphClientFactory){}
   
    async GetJoinedTeams():Promise<Team[]>{
        const graphClient = await this.graphClientFactory.getClient();
        const response = await graphClient.api('me/joinedTeams').get();
        return response.value as Team[];
    }

}