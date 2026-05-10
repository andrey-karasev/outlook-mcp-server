import type { OutlookConfig } from "./config.js";
import type { EmailDetail, EmailSummary, AttachmentSummary } from "./types/email.js";
import type { EmailClient } from "./clients/client.js";
import { GraphEmailClient } from "./clients/graphClient.js";

export class OutlookService {
  private readonly graphClient: EmailClient;

  constructor(private readonly config: OutlookConfig) {
    this.graphClient = new GraphEmailClient(config);
  }

  listEmails(limit: number, folder?: string): Promise<EmailSummary[]> {
    return this.graphClient.listEmails(
      Math.min(limit, this.config.batchSize),
      folder || this.config.mailbox
    );
  }

  getEmail(id: string, markAsRead: boolean = false): Promise<EmailDetail> {
    return this.graphClient.getEmail(id, markAsRead);
  }

  deleteEmail(id: string): Promise<void> {
    return this.graphClient.deleteEmail(id);
  }

  markAsRead(id: string, isRead: boolean): Promise<void> {
    return this.graphClient.markAsRead(id, isRead);
  }

  moveEmail(id: string, destinationFolder: string): Promise<void> {
    return this.graphClient.moveEmail(id, destinationFolder);
  }

  getAttachments(messageId: string): Promise<AttachmentSummary[]> {
    return this.graphClient.getAttachments(messageId);
  }

  getAttachmentContent(messageId: string, attachmentId: string) {
    return this.graphClient.getAttachmentContent(messageId, attachmentId);
  }
}
