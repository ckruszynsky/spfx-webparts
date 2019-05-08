import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { FluentCustomizations } from '@uifabric/fluent-theme';
import { Customizer } from 'office-ui-fabric-react';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';

import { UserProfileCard } from '../UserProfileCard/UserProfileCard';

type MsGraphUserProfileProps = {
  userProfile?: MicrosoftGraph.User;
  photo?: string;  
  personaSize: PersonaSize;
};


type MsGraphUserProfileState = {
  userProfile?: MicrosoftGraph.User;
  photo?: string;
};


export default class MsGraphUserProfile extends React.Component<MsGraphUserProfileProps,MsGraphUserProfileState> {  

  public static defaultProps = {
    personaSize : PersonaSize.size48
  };

  public render(): React.ReactElement<MsGraphUserProfileProps> {
    const {userProfile, photo, personaSize} = this.props;
    
     let userProfileCard =  ( userProfile &&
      <UserProfileCard
        name={userProfile.displayName}
        jobTitle={userProfile.jobTitle}
        emailAddress={userProfile.mail}
        phoneNumber={userProfile.businessPhones.length >0 ? userProfile.businessPhones[0] : ''}
        size={personaSize}
        photoUrl={photo}
     />);
    
    return (
      <div>
      <Customizer {...FluentCustomizations}>      
        {userProfileCard}
      </Customizer>
      </div>
    );
  }
}
