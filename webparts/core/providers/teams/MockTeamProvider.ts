import { Team } from './models/Team';
import { ITeamProvider } from './TeamProvider';


export type MockTeamProviderData = {
    teams: Team[];
}

export class MockTeamProvider implements ITeamProvider {
    constructor(private mockData:MockTeamProviderData){}

    async GetJoinedTeams():Promise<Team[]>{
        return this.mockData.teams;
    }
}