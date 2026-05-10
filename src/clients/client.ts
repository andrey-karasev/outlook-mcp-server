import type { EmailDetail, EmailSummary, AttachmentSummary } from "../types/email.js";

export interface EmailClient {
  listEmails(limit: number, folder?: string): Promise<EmailSummary[]>;
  getEmail(id: string, markAsRead?: boolean): Promise<EmailDetail>;
  deleteEmail(id: string): Promise<void>;
  markAsRead(id: string, isRead: boolean): Promise<void>;
  moveEmail(id: string, destinationFolder: string): Promise<void>;
  getAttachments(id: string): Promise<AttachmentSummary[]>;
  getAttachmentContent(messageId: string, attachmentId: string): Promise<EmailAttachment>;
}

export interface EmailAttachment extends AttachmentSummary {
  contentBytes?: string;
}
