import { Web } from '@pnp/sp';

export interface IWebCommandExecutor {
    GetLists: () =>Promise<any[]>;
}

export class WebCommandExecutor implements IWebCommandExecutor {
    
    constructor(private webUrl:string){}

    async GetLists(fields:string[]=["Title", "DefaultViewUrl"]): Promise<any[]> {
        let web = new Web(this.webUrl);        
        let lists = await web.lists.select(...fields).get();
        return lists;
    }

}