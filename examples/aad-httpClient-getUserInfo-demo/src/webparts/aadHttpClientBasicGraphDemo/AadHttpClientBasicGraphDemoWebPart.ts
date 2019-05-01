import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  AadHttpClient,
  AadHttpClientFactory,
  HttpClientResponse
} from '@microsoft/sp-http';


import styles from './AadHttpClientBasicGraphDemoWebPart.module.scss';
import * as strings from 'AadHttpClientBasicGraphDemoWebPartStrings';

export interface IUserItem {
  displayName: string;
  mail: string;
  userPrincipalName: string;
}
export interface IAadHttpClientBasicGraphDemoWebPartProps {
  description: string;
}

export default class AadHttpClientBasicGraphDemoWebPart extends BaseClientSideWebPart<IAadHttpClientBasicGraphDemoWebPartProps> {

  public async render(): Promise<void> {
    let user = await this.getCurrentUser();

    this.domElement.innerHTML = `
      <div class="${ styles.aadHttpClientBasicGraphDemo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${ user.displayName }: ${ user.mail }</p>
              <p class="${ styles.description}"> ${ user.userPrincipalName} </p>
            </div>
          </div>
        </div>
      </div>`;
  }

  public async getCurrentUser(): Promise<IUserItem> {
    try{
      var client =  await this.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
      var httpResponse = await client.get('https://graph.microsoft.com/v1.0/me', AadHttpClient.configurations.v1);
      return httpResponse.json();
    }
    catch(error) {
      console.log(error);
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
