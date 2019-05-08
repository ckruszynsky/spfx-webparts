import { Link } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {
  IPersonaProps,
  IPersonaSharedProps,
  Persona,
  PersonaPresence,
  PersonaSize,
} from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';

type UserProfileCardProps = {
  name: string;
  jobTitle: string;
  emailAddress: string;
  phoneNumber: string;
  size: PersonaSize;
  photoUrl: string;
};

export class UserProfileCard extends React.Component<UserProfileCardProps, {}> {
  public render(): React.ReactElement<UserProfileCardProps> {
    const { name, jobTitle, emailAddress, phoneNumber, size, photoUrl } = this.props;
    const userPersona: IPersonaSharedProps = {
      secondaryText: jobTitle,
      tertiaryText: emailAddress,
      optionalText: phoneNumber,
      size: size,
      presence: PersonaPresence.online
    };
    return (
      <Persona
        {...userPersona}
        text={name}
        onRenderTertiaryText={this._onRenderTertiaryText}
        onRenderSecondaryText={this._onRenderSecondaryText}
        onRenderOptionalText={this._onRenderOptionalText}
        imageUrl={photoUrl}
      />
    );
  }

  private _onRenderTertiaryText = (props: IPersonaProps): React.ReactElement<any> => {
    if (props.tertiaryText) {
      return (
        <div>
          <Icon iconName={"Mail"} className={"ms-JobIconExample"} /> {"  "}
          <Link href={`mailTo:${props.tertiaryText}`}>{props.tertiaryText}</Link>
        </div>
      );
    }
  };
  private _onRenderOptionalText = (props: IPersonaProps): React.ReactElement<any> => {
    if (props.optionalText) {
      return (
        <div>
          <Icon iconName={"Phone"} className={"ms-JobIconExample"} /> {"  "}
          <Link href={`tel:${props.optionalText}`}>{props.optionalText}</Link>
        </div>
      );
    }
  };
  private _onRenderSecondaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Icon iconName={"Suitcase"} className={"ms-JobIconExample"} />
        {"  "}
        {props.secondaryText}
      </div>
    );
  };
}
