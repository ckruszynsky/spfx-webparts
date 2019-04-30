import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HttpClientExternalOpenApiDemoWebPart.module.scss';
import * as strings from 'HttpClientExternalOpenApiDemoWebPartStrings';

import { IEmployee } from '../../models';
import { EmployeeService } from '../../services';

export interface IHttpClientExternalOpenApiDemoWebPartProps {
  description: string;
}

export default class HttpClientExternalOpenApiDemoWebPart extends BaseClientSideWebPart<IHttpClientExternalOpenApiDemoWebPartProps> {

  //employee service 
  private employeeService: EmployeeService;

  //DOM element where the employee's will go
  private employeeListElement: HTMLElement;

  protected async onInit(): Promise<void> {
    this.employeeService = new EmployeeService(this.context.httpClient);
    return;
  }

  public render(): void {
    if (!this.renderedOnce) {
      this.domElement.innerHTML = `
        <div class="${styles.httpClientExternalOpenApiDemo}">
          <div class="${styles.container}">
            <div class="${styles.row}">
              <div class="${styles.column}">
                <span class="${styles.title}">Employee Directory</span>
                  <p class="${styles.subTitle}">HttpClient External API Demo</p>
                  <button class="${styles.button} getEmployees">Get Employees </button>
                  <button class="${styles.button} createEmployee">Create Employee </button>
                  <div class="employees"></div>
              </div>
            </div>
          </div>
        </div>`;

      // get reference to display HTML element
      this.employeeListElement = document.getElementsByClassName('employees')[0] as HTMLElement;

      //attach event handlers
      this.domElement.getElementsByClassName('getEmployees')[0]
        .addEventListener('click', () => {
          this._getEmployees();
        });

      this.domElement.getElementsByClassName('createEmployee')[0]
        .addEventListener('click', () => {
          this._createEmployee();
        });
    }
  }

  /**
   * Renders collection of employees into the HTML Element specified
   * @private
   * @param {HTMLElement}           element       HTML element to render employees in
   * @param {IEmployee[]}   employees             Collection of employees to render
   * @memberof HttpClientExternalOpenApiDemoWebPart
   */
  private _renderEmployees(element: HTMLElement, employees: IEmployee[]): void {
    let employeeList: string = '';
    if (employees && employees.length && employees.length > 0) {
      employees.forEach((employee: IEmployee) => {
        employeeList = employeeList + `<tr>
            <td>${employee.id}</td>
            <td>${employee.employee_name}</td>
            <td>${employee.employee_age}</td>
            <td>${employee.employee_salary}</td>
            </tr>`;
      });
    }

    element.innerHTML = `<table border=1>
      <tr>
        <th>ID</th>
        <th>Name</th>
        <th>Age</th>
        <th>Salary</th>
      </tr>
      <tbody>${employeeList}</tbody>
      </table>`;
  }

  /**
   * Get all employees
   */
  private async _getEmployees(): Promise<void> {
    try {

      this.context.statusRenderer.displayLoadingIndicator(this.employeeListElement, 'Loading Employees...');
      var employees = await this.employeeService.getEmployees();
      this.context.statusRenderer.clearLoadingIndicator(this.employeeListElement);
      this._renderEmployees(this.employeeListElement, employees);
      return;

    } catch (error) {
      this.context.statusRenderer.clearLoadingIndicator(this.employeeListElement);
      this.context.statusRenderer.renderError(this.employeeListElement, `Error getting Employees : ${error.message}`);
      return;
    }
  }

  /**
   * Create Employee
   */
  private async _createEmployee(): Promise<void> {
    try {
      const newEmployee: IEmployee = {
        id: '6000',
        employee_name: 'Foo',
        employee_age: '21',
        employee_salary: '21',
        profiile_image: ''
      };

      this.context.statusRenderer.displayLoadingIndicator(this.employeeListElement, 'Creating & Loading Employees');

      var employees = await this.employeeService.addEmployee(newEmployee);

      this.context.statusRenderer.clearLoadingIndicator(this.employeeListElement);
      this._renderEmployees(this.employeeListElement,employees);
      return;
    }
    catch (error) {
      this.context.statusRenderer.clearLoadingIndicator(this.employeeListElement);
      this.context.statusRenderer.renderError(this.employeeListElement, `Error saving Employee : ${error.message}`);
    }
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
