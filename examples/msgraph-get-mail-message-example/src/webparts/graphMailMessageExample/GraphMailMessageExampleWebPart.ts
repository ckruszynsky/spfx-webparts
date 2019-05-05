import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'GraphMailMessageExampleWebPartStrings';
import GraphMailMessageExample from './components/GraphMailMessageExample';
import { IGraphMailMessageExampleProps } from './components/IGraphMailMessageExampleProps';
import { MailService } from '../../services';

export interface IGraphMailMessageExampleWebPartProps {
  description: string;
}

export default class GraphMailMessageExampleWebPart extends BaseClientSideWebPart<IGraphMailMessageExampleWebPartProps> {

  public render(): void {
    const service = new MailService(this.context.msGraphClientFactory);    
    const element: React.ReactElement<IGraphMailMessageExampleProps > = React.createElement(
      GraphMailMessageExample,
      {
        mailService:service
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
