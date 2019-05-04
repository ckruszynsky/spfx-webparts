import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'MsGraphUserProfileWebPartStrings';
import MsGraphUserProfile from './components/MsGraphUserProfile';
import { IMsGraphUserProfileProps } from './components/IMsGraphUserProfileProps';
import { GraphService } from '../../services/graphService';

export interface IMsGraphUserProfileWebPartProps {  
}

export default class MsGraphUserProfileWebPart extends BaseClientSideWebPart<IMsGraphUserProfileWebPartProps> {
  
  public render(): void {
    let graphService = new GraphService(this.context.msGraphClientFactory);
    const element: React.ReactElement<IMsGraphUserProfileProps > = React.createElement(
      MsGraphUserProfile,
      {        
        graphService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
