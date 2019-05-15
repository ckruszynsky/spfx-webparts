import { UserProfile } from '../../../models';

export interface IUsersService {
    getCurrentUser: () => Promise<UserProfile>;
    getUserPhoto: (username:string) => Promise<string>;
}