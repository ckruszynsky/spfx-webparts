import * as React from 'react';

import { IMsGraphUserProfileProps } from './IMsGraphUserProfileProps';
import { UserProfileCard } from '../UserProfileCard/UserProfileCard';
import { IMsGraphUserProfileState } from './IMsGraphUserProfileState';
import { Customizer } from 'office-ui-fabric-react';
import { FluentCustomizations } from '@uifabric/fluent-theme';

export default class MsGraphUserProfile extends React.Component<
  IMsGraphUserProfileProps,
  IMsGraphUserProfileState
> {
  constructor(props) {
    super(props);
    this.state = {
      userProfile: null,
      photo: ""
    };
  }

  public async componentDidMount(): Promise<void> {
    const userProfile = await this.props.graphService.getUserProfile();

    this.setState({
      userProfile: userProfile.profile,
      photo: userProfile.photo
    });
    return;
  }

  public render(): React.ReactElement<IMsGraphUserProfileProps> {
    const {userProfile, photo} = this.state;
    let photoUrl = this.props.showUserProfilePhoto ? photo : "";  
    
    return (
      <Customizer {...FluentCustomizations}>
      <div>
        { userProfile &&
        <UserProfileCard
          name={userProfile.displayName}
          jobTitle={userProfile.jobTitle}
          emailAddress={userProfile.mail}
          phoneNumber={userProfile.phone}
          size={this.props.personaSize}
          photoUrl={photoUrl}
        />}
      </div>
      </Customizer>
    );
  }
}
