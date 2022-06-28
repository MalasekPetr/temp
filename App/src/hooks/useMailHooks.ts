import { IFile } from "@pnp/sp/files/types";
import { IItem, IItemAddResult, IItems } from "@pnp/sp/items";
import { IList, List } from "@pnp/sp/lists";
import { Web } from "@pnp/sp/webs";
import { IMailBoxApp, IMessage } from "../models";
import { ConfigurationService, MailService } from "../services";
import "@pnp/sp/files/web";
import { isUndefined, isNull } from 'lodash';
import { replaceElement } from "office-ui-fabric-react";

const RFC2822 = RegExp(/[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?/g)
  
// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types
export const useMailHooks = () => {

    async function saveMessage(spWebBaseUrl: string, spListId: string, m: IMessage, threadId: number, predecessorId: number): Promise<void> {
        // console.log(m);
        const list: IList = List(Web(spWebBaseUrl).lists.getById(spListId));
        const newmail: Promise<IItemAddResult> = list.items.add({
            Title: m.subject,
            MessageThreadId: threadId,
            MessagePredecessorId: predecessorId,
            MessageSubject: m.subject,
            MessageTo: m.to,
            MessageCc: m.cc,
            MessageBcc: m.bcc,
            MessageDate: m.date,
            MessageFrom: m.from,
            MessageSender: isNull(m.sender) ? '' : m.sender,
            MessageReplyTo: m.replyTo,
            MessageTextBody: m.textBody,
            MessageHtmlBody: m.htmlBody,
            MessageHeaders: isNull(m.headers) ? '' : JSON.stringify(m.headers),
            MessageReferences: isNull(m.references) ? '' : JSON.stringify(m.references),
            MessageImportance: isUndefined(m.messageImportance) ? '' : m.messageImportance.toString(),
            MessagePriority: isUndefined(m.priority) ? '' : m.priority.toString(),
            MessageInReplyTo: isNull(m.inReplyTo) ? '' : m.inReplyTo,
            MessageMessageId: m.messageId,
            MessageMimeVersion: isNull(m.mimeVersion) ? '' : JSON.stringify(m.mimeVersion)
        });
        await newmail;
    }

    function listDate(source: string): { value: number; dateFormatted: string } {
        const date: Date = new Date(source);
        return {
          value: date.valueOf(),
          dateFormatted: date.toLocaleDateString(),
        };
      }

    function convertIItems(messages: IItems[]): IMessage[] {
        const result: IMessage[] = [];
        messages.forEach((m: IItems, i: number) => {
            result.push({
                id: m['Id'],
                icon: isNull(m['MessageFrom']) ? '<-' : '->',
                from: isNull(m['MessageFrom']) ? m['MessageSender'] : m['MessageFrom'],
                sender: isNull(m['MessageSender']) ? undefined : m['MessageSender'],
                to: isNull(m['MessageTo']) ? undefined : m['MessageTo'],
                cc: isNull(m['MessageCc']) ? undefined : m['MessageCc'],
                date: isNull(m['MessageDate']) ? undefined : m['MessageDate'],
                htmlBody: isNull(m['MessageHtmlBody']) ? undefined : m['MessageHtmlBody'],
                textBody: isNull(m['MessageTextBody']) ? undefined : m['MessageTextBody'],
                messageId: isNull(m['MessageMessageId']) ? undefined : m['MessageMessageId'],
                inReplyTo: isNull(m['MessageInReplyTo']) ? undefined : m['MessageInReplyTo'],
                bcc: isNull(m['MessageBcc']) ? undefined : m['MessageBcc'],
                threadId: m['MessageThreadId'],
                predecessorId: m['MessagePredecessorId'],
                subject: isNull(m['MessageSubject']) ? m['Title'] : m['MessageSubject'],
                replyTo: isNull(m['MessageReplyTo']) ? undefined : m['MessageReplyTo'],
                mimeVersion: isNull(m['MessageMimeVersion']) ? undefined : m['MessageMimeVersion'],
                headers: isNull(m['MessageHeaders']) ? undefined : m['MessageHeaders'],
                messageImportance: isNull(m['MessageMessageImportance']) ? undefined : m['MessageMessageImportance'],
                priority: isNull(m['MessagePriority']) ? undefined : m['MessagePriority'],
            });
        });
        //console.log(result);
        return result;
    }

/*     async function getMailboxLists(backEndApi: string) : Promise<Array<Record<string, string>>> {
        let items: Array<Record<string, string>>
        const configService: ConfigurationService = new ConfigurationService(backEndApi);  
        const apps = await configService.getMailBoxApps() as IMailBoxApp[];
        apps.forEach(async (app: IMailBoxApp, i: number) => {
            const list: IList = await List(Web(app.spWebBaseUrl).lists.getById(app.spDocLibId).select('Title, Id')).get();
            items.push({key: list['Id'], text: list['Title']})
        });
        // console.log(items);
        return items;
    } */

    async function getAllItems(spWebBaseUrl: string, spListId: string): Promise<IItems[]> {
        const list: IList = List(Web(spWebBaseUrl).lists.getById(spListId));
        const itemsResult: IItems[] = await list.items
            .orderBy(`Id`, false)
            .filter(`MessagePredecessorId eq null`)
            //.select(`MessageSubject, Id, ThreadId, MessageTo`) // TODO: select what is just needed
            .get();
        return itemsResult;
    }

    async function getThreadItems(spWebBaseUrl: string, spListId: string, threadId: number): Promise<IItems[]> {
        const list: IList = List(Web(spWebBaseUrl).lists.getById(spListId));
        const itemsResult: IItems[] = await list.items
            .orderBy(`Id`, false)
            .filter(`MessageThreadId eq '${threadId}'`)
            //.select(`MessageSubject, Id, ThreadId, MessageTo`) // TODO: select what is just needed
            .get();
        return itemsResult;
    }

    async function getThread(spWebBaseUrl: string, spListId: string, threadId: number): Promise<string> {
        const list: IList = List(Web(spWebBaseUrl).lists.getById(spListId));
        const itemsResult: IItems[] = await list.items
            .orderBy(`Id`, false)
            .filter(`(MessageThreadId eq'${threadId}') and (MessagePredecessorId eq null)`)
            .select(`MessageSubject`) // TODO: select what is just needed
            .get();
        return itemsResult.length > 0 ? itemsResult[0]['MessageSubject'] : undefined;
    }

    async function countNewItems(app: IMailBoxApp): Promise<number> {
        const mailService: MailService = new MailService(app.backendapi);
        return await mailService.tryUpdateMailbox(app);
    }

    function convertAddress(sourceAddress: string): string {
        const db = /["]/g;
        const result: string = convertHtml(sourceAddress.replace(db,'\\u0022'));
        return result;
    }    

    function convertVersion(sourceVersion: string): string {
        const ver = /([{"[a-z,A-Z]*":)/g;
        const result: string = sourceVersion.replace(ver,'.')
            .replace('}','') // Removal of tailing curly bracket
            .substring(1); // Removal of leading dot
        return result;
    }    
    
    function removeRe(sourceSubject: string): string {
        const result: string = sourceSubject.replace('RE: ','');
        return result;
    }

    function convertHtml(sourceHtml: string): string {
        const lt = /[<]/g;
        const gt = /[>]/g;
        const result: string = sourceHtml
            .replace(lt,'\\u003C')
            .replace(gt,'\\u003E');
        return result;
    }

    async function sendAutoResponse(backendapi: string, spWebBaseUrl: string, spListId: string, message: IMessage): Promise<boolean> {
        const now = new Date();
        const mailService: MailService = new MailService(backendapi);
        const responseJson = `{
            "to": "${convertAddress(message.from)}",
            "sender": "${convertAddress(message.to)}",
            "subject": "RE: ${removeRe(message.subject)}",
            "htmlBody": "${convertHtml('<html><head></head><body>:-)</body></html>')}",
            "inReplyTo": "${message.messageId}",
            "date": "${now.toISOString()}",
            "priority": 1,
            "messageId": ""
        }`;
        // console.log(message);
        // console.log(responseJson);
        const response: IMessage = await mailService.sendMail(responseJson);
        await saveMessage(spWebBaseUrl, spListId, response, message.threadId, message.id);
        return true;
    }

    function sendFullResponse(backendapi: string, message: IMessage): boolean {
        const mailService: MailService = new MailService(backendapi);
        const messageJson = `{
            "to": "${convertAddress(message.to)}",
            "sender": "${convertAddress(message.sender)}",
            "subject": "${message.subject}",
            "htmlBody": "${convertHtml(message.htmlBody)}",
            "inReplyTo": "${message.messageId}",
            "messageId": ""
        }`;
        // TODO: Add more options (Priority, Attachments, BodyParts ...)
        console.log(message);
        console.log(messageJson);
        mailService.sendMail(messageJson);
        return true;
    }

    async function getMailBoxApps(backEndApi: string): Promise<IMailBoxApp[]> {
        const configService: ConfigurationService = new ConfigurationService(backEndApi);  
        return await configService.getMailBoxApps() as IMailBoxApp[];
    }

    // Return functions
    return {
        getThread,
        getThreadItems,
        sendFullResponse,
        sendAutoResponse,
        convertIItems,
        getAllItems,
        countNewItems,
        getMailBoxApps
    };
};
  