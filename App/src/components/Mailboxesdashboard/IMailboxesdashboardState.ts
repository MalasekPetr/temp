import { IBaseComponentState, IMailBoxApp } from '../../models';

export interface IMailboxesdashboardState extends IBaseComponentState {
    mailboxapps: IMailBoxApp[];
}
