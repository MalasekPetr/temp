import { IBaseComponentState, IMessage } from "../../models";

export interface IMailboxState extends IBaseComponentState {
    items?: IMessage[];
    selectionDetails: string; 
}