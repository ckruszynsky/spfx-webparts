import * as React from 'react';
import styles from './MsGraphUserProfile.module.scss';
import { IMsGraphUserProfileProps } from './IMsGraphUserProfileProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {
  IPersonaSharedProps,
  Persona,
  PersonaInitialsColor,
  PersonaSize,
  PersonaPresence,
  IPersonaProps,
} from 'office-ui-fabric-react/lib/Persona';

export interface IMsGraphUserProfileState {
  userProfile: any;
  photo: any;
}

export default class MsGraphUserProfile extends React.Component<IMsGraphUserProfileProps, IMsGraphUserProfileState> {

  constructor(props) {
    super(props);
    this.state = {
      userProfile: null,
      photo: null
    };
  }

  public async componentDidMount(): Promise<void> {
    console.log('fetching user profile');
    const userProfile = await this.props.graphService.getUserProfile();

    this.setState({
      userProfile: userProfile.profile,
      photo: userProfile.photo
    });
    return;
  }

  public render(): React.ReactElement<IMsGraphUserProfileProps> {
    let details: JSX.Element = <div></div>;
    if (this.state.userProfile) {      
      const userPersona: IPersonaSharedProps = {
        secondaryText: this.state.userProfile.jobTitle,
        tertiaryText: this.state.userProfile.mail,
        optionalText: this.state.userProfile.businessPhone,
        size: PersonaSize.size72,
        presence: PersonaPresence.online
      };
      details = <Persona {...userPersona}
        text={this.state.userProfile.displayName}
        onRenderSecondaryText={this._onRenderSecondaryText}
        imageUrl={this.state.photo}
      />;
    }
    return (
      <div className={styles.msGraphUserProfile}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {details}
            </div>
          </div>
        </div>
      </div>
    );

  }

  private _onRenderSecondaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Icon iconName={'Suitcase'} className={'ms-JobIconExample'} />
        {'  '}{props.secondaryText}
      </div>
    );
  }
}

