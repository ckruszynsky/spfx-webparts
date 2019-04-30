import {
    HttpClient,
    HttpClientResponse
} from '@microsoft/sp-http';

import {IEmployee} from '../models';
const EMPLOYEE_DIRECTORY_ENDPOINT_URI: string = <INSERT ENDPOINT URI>;
export class EmployeeService {
    constructor(private httpClient:HttpClient){}

    public async getEmployees():Promise<IEmployee[]> {
        const endpoint: string = `${EMPLOYEE_DIRECTORY_ENDPOINT_URI}/api/employee-directory`;
        var rawResponse = await this.httpClient.get(endpoint,HttpClient.configurations.v1);
        return rawResponse.json();
    }

    public async addEmployee(employee:IEmployee):Promise<IEmployee[]>{
        const endpoint: string = `${EMPLOYEE_DIRECTORY_ENDPOINT_URI}/api/employee-directory`;
        const request: any = {
            body: JSON.stringify(employee)
        };
        var rawResponse = await this.httpClient.post(endpoint,HttpClient.configurations.v1,request);
        return rawResponse.json();
    }
}