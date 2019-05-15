import { FluentCustomizations } from '@uifabric/fluent-theme';
import { Customizer } from 'office-ui-fabric-react';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';
import { SFC } from 'react';

import { UserProfileCard } from '../../../common/components';
import { UserProfile } from '../../../common/models';

export type UserInformationProps = {
  userProfile?: UserProfile;
  photo?: string;
  personaSize: PersonaSize;
};

const UserInformation: SFC<UserInformationProps> = ({ userProfile, photo, personaSize }) => (
  <>
    <Customizer {...FluentCustomizations}>
      <UserProfileCard
        name={userProfile.displayName}
        jobTitle={userProfile.jobTitle}
        emailAddress={userProfile.emailAddress}
        phoneNumber={userProfile.businessPhone}
        size={personaSize}
        photoUrl={photo}
      />
    </Customizer>
  </>
);

export default UserInformation;
