import { IUser } from './IUser';

export interface IMailBoxApp {
    backendapi?:     string;
    name?:           string;
    address?:        string;
    appAddress?:     string;
    spWebBaseUrl?:   string;
    spDocLibId?:     string;
    spListId?:       string;
    users?:          IUser[];
}