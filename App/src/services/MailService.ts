import {
  HttpClientResponse,
  IHttpClientOptions
} from '@microsoft/sp-http';
import {
  IMailBoxApp,
  IMessage
} from '../models';
import {
  HttpService
} from '.';
import { IList, List } from '@pnp/sp/lists';
import { Web } from '@pnp/sp/webs';
import "@pnp/sp/files/web";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import { isUndefined, isNull } from 'lodash';
import { IItem, IItemAddResult, IItems } from '@pnp/sp/items';
import { IFile, MoveOperations } from '@pnp/sp/files/types';
import { IFolder } from '@pnp/sp/folders';

export class MailService {
  constructor(private endpointUri: string) { }

  public async tryUpdateMailbox(app: IMailBoxApp): Promise<number> {
    const rootFolder: IFolder = await Web(app.spWebBaseUrl).lists.getById(app.spDocLibId).rootFolder.get();
    const files = await Web(app.spWebBaseUrl).getFolderByServerRelativePath(rootFolder['ServerRelativeUrl']).files();
    if (files.length !== 0) {
      await this.updateMailbox(app, rootFolder['ServerRelativeUrl']);
    }
    return files.length;
  }

  public async updateMailbox(app: IMailBoxApp, rootFolder: string): Promise<void> {
    const list: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spListId));
    const doclib: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spDocLibId));
    
    // Find last item in message list with relation to source message file
    const newestItems: IItems = await list.items
      .filter(`MessageThreadId ne null`)
      .orderBy(`Id`, false)
      .top(1)
      .select(`Created, MessageThreadId`)
      .get();
    let lastFiles: IItems;
    if (newestItems.length === 0) {
      // There are no items (first load for fresh mailbox)
      lastFiles = await doclib.items
        .filter(`startswith(ContentTypeId,'0x0101') and substringof('.eml',FileLeafRef)`)
        .select(`Created`)
        .get();
    } else {
      // There are some (standard way)
      lastFiles = await doclib.items
        .filter(`(Id eq ${newestItems[0].MessageThreadId}) and startswith(ContentTypeId,'0x0101') and substringof('.eml',FileLeafRef)`)
        .select(`Created`)
        .get();
    }

    if (lastFiles.length !== 0) {
      await this.organizeNewMessages(app, rootFolder, lastFiles[0].Created)
    }
  }

  public async organizeNewMessages(app: IMailBoxApp, rootFolder: string, lastCreated: string): Promise<void> {
    const doclib: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spDocLibId));

    // Get all source message files from selected date
    const newFiles: IItems[] = await doclib.items
      .filter(`(Created ge '${lastCreated}') and startswith(ContentTypeId,'0x0101') and substringof('.eml',FileLeafRef)`)
      .select(`Id`)
      .expand('File')
      .get();

    // For Each source message files create items in message list
    newFiles.forEach(async (i, id) => {
      const newFile: IItem = await doclib.items
        .getById(i['Id'])
        .select('File/ServerRelativeUrl,File/Name,File/UniqueId,Id')
        .expand('File')
        .get();

      const url = `${newFile['File']['ServerRelativeUrl']}`;
      const proposedUrl = `${rootFolder}/${newFile['File']['Name']}`;    

      if (url === proposedUrl) {
        const fileContent: Promise<string> = Web(app.spWebBaseUrl).getFileByServerRelativeUrl(url).getText();
        const m: IMessage = await this.parseMail(btoa(await fileContent));
        const threadId: number = i['Id']; // ThreadId = MessageThreadId = Id of source incoming message
        if (!isNull(m.inReplyTo)) {
          await this.newReply(app, rootFolder, threadId, newFile, m);
        } else {
          await this.newThread(app, rootFolder, threadId, newFile, m);
        }
      } else {
        // TODO manage orphaned files
      }
    });
  }
  
  public async newThread(app: IMailBoxApp, rootFolder: string, threadId: number, file: IItem, m: IMessage): Promise<void> {
    const list: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spListId));
    const doclib: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spDocLibId));
    const url = `${file['File']['ServerRelativeUrl']}`;

    // Check if there are not related item in the list
    const threadItems: IItem = await list.items
      .filter(`(MessageThreadId eq ${threadId}) and (MessagePredecessorId eq null)`)
      .select(`Id`).top(1).get();
    if (threadItems.length === 0) {
      // Move file to New subfolder
      const folderName = file['File']['UniqueId'];
      const folderAdded = await Web(app.spWebBaseUrl).lists.getById(app.spDocLibId).rootFolder.folders.add(folderName);
      const destinationUrl = file['File']['ServerRelativeUrl'].toString().replace(file['File']['Name'], `${folderName}/${file['File']['Name']}`);
      const eml = await Web(app.spWebBaseUrl)
        .getFileByServerRelativePath(`/${file['File']['ServerRelativeUrl']}`)
        .moveTo(destinationUrl);

      if (m.files) {
        m.files.forEach(async (f, id) => {
          console.log(f);
          const attSourceUrl = `${rootFolder}/${f}`;
          console.log(attSourceUrl);
          const attSource: boolean = await Web(app.spWebBaseUrl).getFileByServerRelativePath(attSourceUrl).exists();
          console.log(attSource);
          if (attSource) {
            // Move attachment file to Thread subfolder
            let attDestinationUrl = attSourceUrl.replace(f, `${folderName}/${f}`);
            console.log(attDestinationUrl);
            const attDestination: boolean = await Web(app.spWebBaseUrl).getFileByServerRelativePath(attDestinationUrl).exists();
            if (attDestination) {
              attDestinationUrl = attDestinationUrl.replace(f, `${new Date().getTime().toString()}-${f}`);
            }
            const att = await Web(app.spWebBaseUrl)
              .getFileByServerRelativePath(attSourceUrl)
              .moveTo(attDestinationUrl);
          } else {
            const allitems: any[] = await doclib.items
            .expand('File')
            .filter(`startswith(ContentTypeId,'0x0101')`)
            .orderBy(`Id`, false)
            .top(1000)
            .select(`EmailHeaders,File/ServerRelativeUrl,File/Name`)
            .get();
            console.log(allitems);
            allitems.forEach(async (i, id) => {
              const headers: string = i['EmailHeaders'];
              const url = `${i['File']['ServerRelativeUrl']}`;
              const proposedUrl = `${rootFolder}/${i['File']['Name']}`;
              if (url === proposedUrl) {
                if (headers) {
                  console.log(headers);
                  console.log(m.headers);
                  console.log(m.messageId);
                  if (headers.lastIndexOf(m.messageId) > 0) {
                    const attAltSource: boolean = await Web(app.spWebBaseUrl)
                      .getFileByServerRelativePath(i['File']['ServerRelativeUrl']).exists();
                    if (attAltSource) {
                      // Move attachment file to Thread subfolder
                      let attAltDestinationUrl = attSourceUrl.replace(i['File']['ServerRelativeUrl'], `${folderName}/${i['File']['ServerRelativeUrl']}`);
                      const attAltDestination: boolean = await Web(app.spWebBaseUrl).getFileByServerRelativePath(attAltDestinationUrl).exists();
                      if (attAltDestination) {
                        attAltDestinationUrl = attAltDestinationUrl.replace(f, `${new Date().getTime().toString()}-${f}`);
                      }
                      const altAtt = await Web(app.spWebBaseUrl)
                        .getFileByServerRelativePath(attSourceUrl)
                        .moveTo(attAltDestinationUrl);   
                    }    
                  }
                }
              }
            });
          }
        });
      }

      // Save new Thread message
      const newmail: Promise<IItemAddResult> = list.items.add({
        Title: isNull(m.subject) ? file['File']['Name'] : m.subject,
        MessageThreadId: threadId,
        MessagePredecessorId: null,
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
  }

  public async newReply(app: IMailBoxApp, rootFolder: string, threadId: number, file: IItem, m: IMessage): Promise<void> {
    const list: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spListId));
    const doclib: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spDocLibId));
    let predecessorId: number = null;
    const url = `${file['File']['ServerRelativeUrl']}`;

    // The message is Response, so try to find source
    const prevItem: IItem = await list.items
      .filter(`MessageMessageId eq '${m.inReplyTo}'`)
      /* .select(`Id, MessageThreadId`) */ // TODO: uncomment for production
      .top(1)
      .get();
    if ((prevItem.length !== 0)) {

      predecessorId = prevItem[0]['Id']; // PredecessorId = MessagePredecessorId = Id of message in reply to
      threadId = prevItem[0]['MessageThreadId'];

      // Check if the message is not in the list, yet
      const duplicity: IItem = await list.items
        .filter(`(MessageThreadId eq ${threadId}) and (MessagePredecessorId eq ${predecessorId}) and (MessageMessageId eq '${m.messageId}')`)
        .select(`Id`).top(1).get();
      if (duplicity.length !== 0) {
        return // Nothing more to do
      }

      const prevFile: IItem = await doclib.items
        .getById(threadId)
        /* .select('File/ServerRelativeUrl, Id') */ // TODO: uncomment for production
        .expand('File')
        .get();

      // Move file to Thread subfolder
      let destinationUrl = prevFile['File']['ServerRelativeUrl'].toString().replace(prevFile['File']['Name'], file['File']['Name']);
      const fileExists: boolean = await Web(app.spWebBaseUrl).getFileByServerRelativePath(destinationUrl).exists();
      if (fileExists) {
        destinationUrl = destinationUrl.replace('.eml', `-${new Date().getTime().toString()}.eml`);
      }
      await Web(app.spWebBaseUrl)
        .getFileByServerRelativePath(`/${file['File']['ServerRelativeUrl']}`)
        .moveTo(destinationUrl);

      if (m.files) {
        m.files.forEach(async (f, id) => {
          const attUrl = url.replace(file['File']['Name'], f);
          console.log(attUrl);
          /* const att: boolean = await Web(app.spWebBaseUrl).getFileByServerRelativePath(url).exists();
          if (att) {
            // Move attachment file to Thread subfolder
            const attExists: boolean = await Web(app.spWebBaseUrl).getFileByServerRelativePath(url).exists();
            if (attExists) {
              destinationUrl = destinationUrl.replace('.eml', `-${new Date().getTime().toString()}.eml`);
            }
            await Web(app.spWebBaseUrl)
              .getFileByServerRelativePath(`/${newFile['File']['ServerRelativeUrl']}`)
              .moveTo(destinationUrl);
          } */
        });
      }

      // Save new Response as part of the Thread message
      const newmail: Promise<IItemAddResult> = list.items.add({
        Title: isNull(m.subject) ? file['File']['Name'] : m.subject,
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
  }

  public async getMessages(app: IMailBoxApp): Promise<IItems[]> {
    return new Promise<IItems[]>((resolve, reject) => {
      const list: IList = List(Web(app.spWebBaseUrl).lists.getById(app.spListId));
      list.items
        .orderBy(`Id`, false)
        .select(`MessageSubject`)
        .get()
        .then((res: IItems[]) => {
          resolve(res);
        })
        .catch((err: Record<string, unknown>) => {
          const response: Record<string, unknown> = err;
          console.error(err);
          reject(response);
        });
    }
    )
  }

  public async parseMail(base64content: string): Promise<IMessage> {
    return new Promise<IMessage>((resolve, reject) => {
      const endpoint = `${this.endpointUri}/api/message/parse`;
      const headers: Headers = new Headers({
        'CONTENT-TYPE': 'application/json'
      });
      let options: IHttpClientOptions = { headers: headers, mode: 'cors' };
      options = { headers: headers, mode: 'cors', body: `{"msgcontent":"${base64content}"}` };
      HttpService.Post(endpoint, options)
        .then((rawResponse: HttpClientResponse) => {
          return rawResponse.json();
        })
        .then((jsonResponse: IMessage) => {
          resolve(jsonResponse);
        })
        .catch((err: Record<string, unknown>) => {
          const response: Record<string, unknown> = err;
          console.error(err);
          reject(response);
        });
    });
  }

  public async sendMail(message: string): Promise<IMessage> {
    return new Promise<IMessage>((resolve, reject) => {
      const endpoint = `${this.endpointUri}/api/message/send`;
      const headers: Headers = new Headers({
        'CONTENT-TYPE': 'application/json'
      });
      let options: IHttpClientOptions = { headers: headers, mode: 'cors' };
      options = { headers: headers, mode: 'cors', body: message };
      HttpService.Post(endpoint, options)
        .then((rawResponse: HttpClientResponse) => {
          return rawResponse.json();
        })
        .then((jsonResponse: IMessage) => {
          resolve(jsonResponse);
        })
        .catch((err: Record<string, unknown>) => {
          const response: Record<string, unknown> = err;
          console.error(err);
          reject(response);
        });
    });
  }
}