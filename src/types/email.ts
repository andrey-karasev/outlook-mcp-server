export interface EmailSummary {
  id: string;
  subject: string;
  from: string;
  fromAddress?: string;
  date: string | null;
  isRead: boolean;
  snippet: string;
  receivedDateTime: string;
}

export interface EmailDetail extends EmailSummary {
  to: string;
  toAddresses?: string[];
  cc?: string;
  ccAddresses?: string[];
  bcc?: string;
  bccAddresses?: string[];
  bodyPreview: string;
  body: string;
  bodyType: "text" | "html";
  isReminderOn?: boolean;
  isFlagged?: boolean;
  importance: "low" | "normal" | "high";
  hasAttachments: boolean;
}

export interface AttachmentSummary {
  id: string;
  name: string;
  size: number;
  mimeType: string;
}

export interface EmailAttachment extends AttachmentSummary {
  contentBytes?: string;
}
