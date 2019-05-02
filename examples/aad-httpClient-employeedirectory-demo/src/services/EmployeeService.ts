import {
    AadHttpClient,
    AadHttpClientFactory,
    HttpClientResponse
} from '@microsoft/sp-http';


import {IEmployee} from '../models';


const EMPLOYEE_DIRECTORY_ENDPOINT_URI: string ="<your endpoint here>";
export class EmployeeService {
    constructor(private aadHttpClientFactory:AadHttpClientFactory){}

    public async getEmployees():Promise<IEmployee[]> {
        const endpoint: string = `${EMPLOYEE_DIRECTORY_ENDPOINT_URI}<your api here>`;
        var client = await this.aadHttpClientFactory.getClient(EMPLOYEE_DIRECTORY_ENDPOINT_URI);   
        var httpResponse =await client.get(endpoint,AadHttpClient.configurations.v1);

         // verify successful response
         if (httpResponse.status === 200) {
            return httpResponse.json();
          } else {
            throw new Error('Error occurred when retrieving employees.');
          }        
    }

    public async addEmployee(employee:IEmployee):Promise<IEmployee[]>{
        const endpoint: string = `${EMPLOYEE_DIRECTORY_ENDPOINT_URI}<your api here>`;
        const request: any = {
            body: JSON.stringify(employee)
        };
        var client = await this.aadHttpClientFactory.getClient(EMPLOYEE_DIRECTORY_ENDPOINT_URI);   
        var httpResponse =await client.post(endpoint,AadHttpClient.configurations.v1,request);
         // verify successful response
         if (httpResponse.status === 200) {
            return httpResponse.json();
          } else {
            throw new Error('Error occurred when retrieving employees.');
          }   
    }
}