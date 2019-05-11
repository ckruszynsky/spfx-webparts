import { IWebCommandExecutor } from './WebCommandExecutor';


export type MockWebCommandExecutorData = {
    lists: any[];
}
export class MockWebCommandExecutor implements IWebCommandExecutor {
    constructor(private mockData:MockWebCommandExecutorData){}
    
    async GetLists(): Promise<any[]> {
        return this.mockData.lists;
    }

}