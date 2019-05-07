import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IPersonaProps, IPersonaSharedProps, Persona, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { IUserProfileCardProps } from "./IUserProfileCardProps";
import * as React from 'react';

export class UserProfileCard extends React.Component<IUserProfileCardProps, {}> {
  public render(): React.ReactElement<IUserProfileCardProps> {
    const { name, jobTitle, emailAddress, phoneNumber, size, photoUrl } = this.props;
    const userPersona: IPersonaSharedProps = {
      secondaryText: jobTitle,
      tertiaryText: emailAddress,
      optionalText: phoneNumber,
      size: size,
      presence: PersonaPresence.online
    };
    return (<Persona {...userPersona} text={name} onRenderSecondaryText={this._onRenderSecondaryText} imageUrl={photoUrl} />);
  }
  private _onRenderSecondaryText = (props: IPersonaProps): JSX.Element => {
    return (<div>
      <Icon iconName={"Suitcase"} className={"ms-JobIconExample"} />
      {"  "}
      {props.secondaryText}
    </div>);
  };
}
