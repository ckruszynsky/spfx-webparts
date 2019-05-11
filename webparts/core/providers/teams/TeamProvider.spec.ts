import 'jest';

import { MockTeamCommandExecutor } from './executor/MockTeamCommandExecutor';
import TeamProvider from './TeamProvider';

describe("Provider: Team", () => {

    
  it("should get a team", async () => {
    //arrange
    const provider = new TeamProvider(
      new MockTeamCommandExecutor({
        teams: [
          {
            id:"1",
            displayName:"Mock Team",
            description:"A mocked team",
            isArchived:false
          }
        ]
      })
    );

    //act
    const teams = await provider.GetJoinedTeams();

    //assert
    expect(teams.length).toBe(1);
  });
})
