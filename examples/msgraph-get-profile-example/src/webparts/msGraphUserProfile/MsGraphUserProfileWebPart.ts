import { Environment, EnvironmentType, Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart, PropertyPaneDropdown } from "@microsoft/sp-webpart-base";
import * as strings from "MsGraphUserProfileWebPartStrings";
import { PersonaSize } from "office-ui-fabric-react/lib/Persona";
import * as React from "react";
import * as ReactDom from "react-dom";

import { GraphService } from "../../services/graphService";
import { IMsGraphUserProfileProps } from "./components/MsGraphUserProfile/IMsGraphUserProfileProps";
import MsGraphUserProfile from "./components/MsGraphUserProfile/MsGraphUserProfile";
import { CalloutTriggers } from "@pnp/spfx-property-controls/lib/PropertyFieldHeader";
import { PropertyFieldToggleWithCallout } from "@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout";

export interface IMsGraphUserProfileWebPartProps {
  showUserPhoto: boolean;
  personaSize: PersonaSize;
}

export default class MsGraphUserProfileWebPart extends BaseClientSideWebPart<
  IMsGraphUserProfileWebPartProps
> {
  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }
  public render(): void {
    const userProfileProps: IMsGraphUserProfileProps = {
      graphService: new GraphService(this.context.msGraphClientFactory),
      showUserProfilePhoto: this.properties.showUserPhoto,
      personaSize: this.properties.personaSize
    };

    const element: React.ReactElement<IMsGraphUserProfileProps> = React.createElement(
      MsGraphUserProfile,
      userProfileProps
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
              groupFields: [
                PropertyFieldToggleWithCallout(strings.ShowPhotoTargetProperty, {
                  calloutTrigger: CalloutTriggers.Click,
                  key: strings.ShowPhotoTargetProperty,
                  label: strings.ShowPhotoLabel,
                  calloutContent: React.createElement(
                    "p",
                    {},
                    strings.ShowPhotoCalloutText
                  ),
                  onText: strings.ShowPhotoOnText,
                  offText: strings.ShowPhotoOffText,
                  checked: this.properties.showUserPhoto
                }),
                PropertyPaneDropdown(strings.PersonaSizeTargetProperty, {
                  label: strings.PersonaSizeLabel,
                  options: [
                    { key: PersonaSize.size24, text: "Extra Small" },
                    { key: PersonaSize.size40, text: "Small" },
                    { key: PersonaSize.size48, text: "Normal" },
                    { key: PersonaSize.size72, text: "Large" },
                    { key: PersonaSize.size100, text: "Extra Large" }
                  ],
                  selectedKey: PersonaSize.size48
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (oldValue !== newValue) {
      this.properties[propertyPath] = newValue;
      this.render();
    }
  }
  /**
   * Provides logic to update web part properties and initiate re-render
   * @param targetProperty property that has been changed
   * @param newValue new value of the property
   */
  public onCustomPropertyPaneFieldChanged(targetProperty: string, newValue: any) {
    const oldValue = this.properties[targetProperty];
    this.properties[targetProperty] = newValue;

    this.onPropertyPaneFieldChanged(targetProperty, oldValue, newValue);

    // NOTE: in local workbench onPropertyPaneFieldChanged method initiates re-render
    // in SharePoint environment we need to call re-render by ourselves
    if (Environment.type !== EnvironmentType.Local) {
      this.render();
    }
  }
}
