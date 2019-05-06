import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'GraphGetUserGroupsExampleWebPartStrings';
import GraphGetUserGroupsExample from './components/GraphGetUserGroupsExample';
import { IGraphGetUserGroupsExampleProps } from './components/IGraphGetUserGroupsExampleProps';
import { GroupService } from '../../services/GroupService';

export interface IGraphGetUserGroupsExampleWebPartProps {
  description: string;
}

export default class GraphGetUserGroupsExampleWebPart extends BaseClientSideWebPart<IGraphGetUserGroupsExampleWebPartProps> {

  public render(): void {
    const service = new GroupService(this.context.msGraphClientFactory);
    const element: React.ReactElement<IGraphGetUserGroupsExampleProps > = React.createElement(
      GraphGetUserGroupsExample,
      {
        service: service
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
