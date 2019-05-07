import { GraphService } from "../../../../services/graphService";
import { PersonaSize } from "office-ui-fabric-react/lib/Persona";

export interface IMsGraphUserProfileProps {
  graphService:GraphService;
  showUserProfilePhoto:boolean;
  personaSize: PersonaSize;
}
