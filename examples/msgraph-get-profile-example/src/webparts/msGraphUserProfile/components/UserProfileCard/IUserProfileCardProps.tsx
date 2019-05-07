import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
export interface IUserProfileCardProps {
  name: string;
  jobTitle: string;
  emailAddress: string;
  phoneNumber: string;
  size: PersonaSize;
  photoUrl: string;
}
