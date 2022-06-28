export interface IMessage {
    id?:                 number;
    icon?:               string;
    messageImportance?:  string;
    priority?:           string;
    sender?:             string;
    from?:               string;
    replyTo?:            string;
    to:                  string;
    cc?:                 string;
    bcc?:                string;
    subject?:            string;
    date?:               string;
    references?:         string;
    inReplyTo?:          string;
    messageId?:          string;
    mimeVersion?:        string;
    textBody?:           string;
    htmlBody?:           string;
    headers?:            string;
    redBy?:              string;
    threadId?:           number;
    predecessorId?:      number;
    attachments?:        string[];
    files?:              string[];
}