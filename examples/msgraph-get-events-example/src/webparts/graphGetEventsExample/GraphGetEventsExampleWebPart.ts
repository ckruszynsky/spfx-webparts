import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'GraphGetEventsExampleWebPartStrings';
import GraphGetEventsExample from './components/GraphGetEventsExample';
import { IGraphGetEventsExampleProps } from './components/IGraphGetEventsExampleProps';
import { EventService } from '../../services/index';

export interface IGraphGetEventsExampleWebPartProps {
  description: string;
}

export default class GraphGetEventsExampleWebPart extends BaseClientSideWebPart<IGraphGetEventsExampleWebPartProps> {

  public render(): void {
    const eventService = new EventService(this.context.msGraphClientFactory);
    const element: React.ReactElement<IGraphGetEventsExampleProps > = React.createElement(
      GraphGetEventsExample,
      {
        eventService: eventService
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
