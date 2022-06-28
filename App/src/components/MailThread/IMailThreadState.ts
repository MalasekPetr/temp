import { IBaseComponentState, IMessage } from "../../models";

export interface IMailThreadState extends IBaseComponentState {
    items?: IMessage[];
    selectionDetails: string; 
}