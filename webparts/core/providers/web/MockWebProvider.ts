import { WebListInformation } from './models/WebListInformation';
import { IWebProvider } from './WebProvider';

export type MockWebProviderData = {
    lists:WebListInformation[];
}
export class MockWebProvider implements IWebProvider {    
    constructor(private mockData:MockWebProviderData){}

    async GetLists():Promise<WebListInformation[]> {
        return this.mockData.lists;
    }
        
    
    
}