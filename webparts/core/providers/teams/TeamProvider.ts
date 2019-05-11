import { ITeamCommandExecutor } from './executor/TeamCommandExecutor';
import { Team } from './models/Team';


export interface ITeamProvider {
    GetJoinedTeams: () => Promise<Team[]>
}

export default class TeamProvider implements ITeamProvider {
    constructor(private executor:ITeamCommandExecutor){}

    async GetJoinedTeams():Promise<Team[]>{
        return await this.executor.GetJoinedTeams();
    }
}