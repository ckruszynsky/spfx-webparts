import { UserProfile } from '../../models';

export interface IUserProfileService {
    getCurrentUser: () => Promise<UserProfile>;
    getUserPhoto: (username:string) => Promise<string>;
}