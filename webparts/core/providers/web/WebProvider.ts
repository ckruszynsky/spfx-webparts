import { IWebCommandExecutor } from './executor/WebCommandExecutor';
import { WebListInformation } from './models/WebListInformation';

export interface IWebProvider {
    GetLists: () => Promise<WebListInformation[]>;
}
export default class WebProvider implements IWebProvider {

    constructor(private executor:IWebCommandExecutor){}

    async GetLists():Promise<WebListInformation[]> {
        const listData = await this.executor.GetLists();
        const lists = listData.map(l => {
            return { 
                Title: l.Title,
                DefaultViewUrl: l.DefaultViewUrl
            }
        });
        return lists;
    }
    
}