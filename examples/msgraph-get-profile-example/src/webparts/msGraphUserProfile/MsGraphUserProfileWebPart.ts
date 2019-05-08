import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import * as strings from 'MsGraphUserProfileWebPartStrings';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { GraphService } from '../../services/graphService';
import MsGraphUserProfile from './components/MsGraphUserProfile/MsGraphUserProfile';

export interface IMsGraphUserProfileWebPartProps {
  showUserPhoto: boolean;
  personaSize: PersonaSize;
}
type MsGraphUserProfileProps = MsGraphUserProfile["props"];

const showPhotoToggleField = PropertyFieldToggleWithCallout(strings.ShowPhotoTargetProperty, {
  calloutTrigger: CalloutTriggers.Click,
  key: strings.ShowPhotoTargetProperty,
  label: strings.ShowPhotoLabel,
  calloutContent: React.createElement("p", {}, strings.ShowPhotoCalloutText),
  onText: strings.ShowPhotoOnText,
  offText: strings.ShowPhotoOffText,
  checked: this.properties.showUserPhoto
});

const personaSizeDropdownField = PropertyPaneDropdown(strings.PersonaSizeTargetProperty, {
  label: strings.PersonaSizeLabel,
  options: [
    { key: PersonaSize.size24, text: "Extra Small" },
    { key: PersonaSize.size40, text: "Small" },
    { key: PersonaSize.size48, text: "Normal" },
    { key: PersonaSize.size72, text: "Large" },
    { key: PersonaSize.size100, text: "Extra Large" }
  ],
  selectedKey: PersonaSize.size48
});

const propertyPaneConfiguration = {
  header: {
    description: strings.PropertyPaneDescription
  },
  groups: [
    {
      groupFields: [showPhotoToggleField, personaSizeDropdownField]
    }
  ]
};

export default class MsGraphUserProfileWebPart extends BaseClientSideWebPart<IMsGraphUserProfileWebPartProps> {
  private _userInformation: MicrosoftGraph.User;
  private _profilePhotoUrl: string = "";

  public async onInit(): Promise<void> {
    let graphService = new GraphService(this.context.msGraphClientFactory);
    this._userInformation = await graphService.getUserProfile();

    if (this.properties.showUserPhoto) {
      this._profilePhotoUrl = graphService.getProfilePhoto(this._userInformation.userPrincipalName);
    }
    return;
  }
  public render(): void {
    const userProfileProps: MsGraphUserProfileProps = {
      userProfile: this._userInformation,
      photo: this._profilePhotoUrl,
      personaSize: this.properties.personaSize
    };

    const element: React.ReactElement<MsGraphUserProfileProps> = React.createElement(
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
      pages: [propertyPaneConfiguration]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (oldValue !== newValue) {
      this.properties[propertyPath] = newValue;
      this.render();
    }
  }
}
