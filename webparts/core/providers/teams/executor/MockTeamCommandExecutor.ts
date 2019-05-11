import { Team } from '../models/Team';
import { ITeamCommandExecutor } from './TeamCommandExecutor';



export type MockTeamCommandExecutorData = {
    teams: Team[];
}
export class MockTeamCommandExecutor implements ITeamCommandExecutor {
    constructor(private mockData:MockTeamCommandExecutorData){}
    
    async GetJoinedTeams(): Promise<Team[]> {
        return this.mockData.teams;
    }

}