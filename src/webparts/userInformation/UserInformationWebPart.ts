import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PersonaSize } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'UserInformationWebPartStrings';

import { UserProfile } from '../../common/models';
import UsersService from '../../common/services/graph/users/UsersService';
import UserInformation, { UserInformationProps } from './components/UserInformation';

export interface IUserInformationWebPartProps {
  showUserPhoto: boolean;
  personaSize: PersonaSize;
}

export default class UserInformationWebPart extends BaseClientSideWebPart<IUserInformationWebPartProps> {
  private currentUser: UserProfile;
  private photoUrl: string;

  public async onInit(): Promise<void> {
    const usersService = this.context.serviceScope.consume(UsersService.serviceKey);
    this.currentUser = await usersService.getCurrentUser();

    if (this.properties.showUserPhoto) {
      this.photoUrl = await usersService.getUserPhoto(this.currentUser.emailAddress);
    }
  }
  public render(): void {
    const userInfoProps: UserInformationProps = {
      userProfile: this.currentUser,
      photo: this.properties.showUserPhoto ? this.photoUrl : '',
      personaSize: this.properties.personaSize
    };

    const element: React.ReactElement<UserInformationProps> = React.createElement(UserInformation, userInfoProps);

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

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log(`Propery Pane Field Changed for Property Path: ${propertyPath} replacing ${oldValue} with ${newValue}`);
    this.properties[propertyPath] = newValue;
    this.render();
  }
}
