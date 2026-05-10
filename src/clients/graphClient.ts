import https from "https";
import { Client } from "@microsoft/microsoft-graph-client";
import type { EmailDetail, EmailSummary, AttachmentSummary } from "../types/email.js";
import type { EmailClient, EmailAttachment } from "./client.js";
import type { OutlookConfig } from "../config.js";

export class GraphEmailClient implements EmailClient {
  private graphClient: Client;
  private readonly config: OutlookConfig;

  constructor(config: OutlookConfig) {
    this.config = config;

    if (!this.config.refreshToken) {
      throw new Error(
        "Missing OUTLOOK_REFRESH_TOKEN in environment. Delegated authentication requires a refresh token."
      );
    }

    this.graphClient = Client.init({
      authProvider: async (done) => {
        try {
          const accessToken = await this.getAccessToken();
          done(null, accessToken);
        } catch (error) {
          done(error, null);
        }
      }
    });
  }

  private getAccessToken(): Promise<string> {
    return new Promise((resolve, reject) => {
      const postData = new URLSearchParams({
        client_id: this.config.clientId,
        client_secret: this.config.clientSecret,
        refresh_token: this.config.refreshToken || "",
        grant_type: "refresh_token",
        scope: "offline_access mail.read mail.readwrite"
      }).toString();

      const options = {
        hostname: "login.microsoftonline.com",
        path: `/${this.config.tenantId}/oauth2/v2.0/token`,
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Content-Length": Buffer.byteLength(postData)
        }
      };

      const req = https.request(options, (res) => {
        let body = "";
        res.on("data", (chunk) => {
          body += chunk;
        });
        res.on("end", () => {
          try {
            const parsed = JSON.parse(body);
            if (res.statusCode === 200 && parsed.access_token) {
              resolve(parsed.access_token);
            } else {
              reject(
                new Error(
                  `Failed to acquire access token: ${parsed.error_description || body}`
                )
              );
            }
          } catch (err) {
            reject(err);
          }
        });
      });

      req.on("error", reject);
      req.write(postData);
      req.end();
    });
  }

  async listEmails(limit: number, folder: string = "inbox"): Promise<EmailSummary[]> {
    try {
      const messages = await this.graphClient
        .api(`/me/mailFolders/${folder}/messages`)
        .select(
          "id,subject,from,receivedDateTime,isRead,bodyPreview,webLink"
        )
        .orderby("receivedDateTime desc")
        .top(Math.min(limit, 999))
        .get();

      return messages.value.map((msg: any) => this.mapToEmailSummary(msg));
    } catch (error) {
      throw new Error(`Failed to list emails: ${error}`);
    }
  }

  async getEmail(id: string, markAsRead: boolean = false): Promise<EmailDetail> {
    try {
      const message = await this.graphClient
        .api(`/me/messages/${id}`)
        .select(
          "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,isRead,bodyPreview,body,importance,hasAttachments,webLink,flag"
        )
        .get();

      if (markAsRead && !message.isRead) {
        await this.markAsRead(id, true);
      }

      return this.mapToEmailDetail(message);
    } catch (error) {
      throw new Error(`Failed to get email ${id}: ${error}`);
    }
  }

  async deleteEmail(id: string): Promise<void> {
    try {
      await this.graphClient.api(`/me/messages/${id}`).delete();
    } catch (error) {
      throw new Error(`Failed to delete email ${id}: ${error}`);
    }
  }

  async markAsRead(id: string, isRead: boolean): Promise<void> {
    try {
      await this.graphClient
        .api(`/me/messages/${id}`)
        .patch({ isRead });
    } catch (error) {
      throw new Error(`Failed to mark email ${id} as read: ${error}`);
    }
  }

  async listFolders(): Promise<{ id: string; displayName: string }[]> {
    try {
      const result = await this.graphClient
        .api("/me/mailFolders")
        .select("id,displayName")
        .top(50)
        .get();
      return result.value.map((f: any) => ({ id: f.id, displayName: f.displayName }));
    } catch (error) {
      throw new Error(`Failed to list folders: ${error}`);
    }
  }

  async moveEmail(id: string, destinationFolder: string): Promise<void> {
    try {
      // Resolve folder name to ID if not already an ID
      let folderId = destinationFolder;
      if (!destinationFolder.match(/^[A-Za-z0-9_=-]{20,}$/)) {
        const folders = await this.listFolders();
        const match = folders.find(
          (f) => f.displayName.toLowerCase() === destinationFolder.toLowerCase()
        );
        if (match) {
          folderId = match.id;
        } else {
          // Create the folder
          const created = await this.graphClient
            .api("/me/mailFolders")
            .post({ displayName: destinationFolder });
          folderId = created.id;
        }
      }
      await this.graphClient.api(`/me/messages/${id}/move`).post({ destinationId: folderId });
    } catch (error) {
      throw new Error(`Failed to move email ${id}: ${error}`);
    }
  }

  async getAttachments(messageId: string): Promise<AttachmentSummary[]> {
    try {
      const attachments = await this.graphClient
        .api(`/me/messages/${messageId}/attachments`)
        .select("id,name,size,@odata.type,mimeType")
        .get();

      return attachments.value
        .filter((att: any) => att["@odata.type"] !== "#microsoft.graph.itemAttachment")
        .map((att: any) => ({
          id: att.id,
          name: att.name,
          size: att.size || 0,
          mimeType: att.mimeType || "application/octet-stream"
        }));
    } catch (error) {
      throw new Error(
        `Failed to get attachments for message ${messageId}: ${error}`
      );
    }
  }

  async getAttachmentContent(
    messageId: string,
    attachmentId: string
  ): Promise<EmailAttachment> {
    try {
      const attachment = await this.graphClient
        .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
        .select("id,name,size,@odata.type,mimeType,contentBytes")
        .get();

      return {
        id: attachment.id,
        name: attachment.name,
        size: attachment.size || 0,
        mimeType: attachment.mimeType || "application/octet-stream",
        contentBytes: attachment.contentBytes
      };
    } catch (error) {
      throw new Error(
        `Failed to get attachment ${attachmentId} content: ${error}`
      );
    }
  }

  private mapToEmailSummary(msg: any): EmailSummary {
    return {
      id: msg.id,
      subject: msg.subject || "(no subject)",
      from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
      fromAddress: msg.from?.emailAddress?.address,
      date: msg.receivedDateTime,
      isRead: msg.isRead,
      snippet: msg.bodyPreview || "",
      receivedDateTime: msg.receivedDateTime
    };
  }

  private mapToEmailDetail(msg: any): EmailDetail {
    const toAddresses = msg.toRecipients?.map((r: any) => r.emailAddress?.address).filter(Boolean) || [];
    const ccAddresses = msg.ccRecipients?.map((r: any) => r.emailAddress?.address).filter(Boolean) || [];
    const bccAddresses = msg.bccRecipients?.map((r: any) => r.emailAddress?.address).filter(Boolean) || [];

    return {
      id: msg.id,
      subject: msg.subject || "(no subject)",
      from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
      fromAddress: msg.from?.emailAddress?.address,
      to: toAddresses.join(", "),
      toAddresses,
      cc: ccAddresses.length > 0 ? ccAddresses.join(", ") : undefined,
      ccAddresses: ccAddresses.length > 0 ? ccAddresses : undefined,
      bcc: bccAddresses.length > 0 ? bccAddresses.join(", ") : undefined,
      bccAddresses: bccAddresses.length > 0 ? bccAddresses : undefined,
      date: msg.receivedDateTime,
      isRead: msg.isRead,
      bodyPreview: msg.bodyPreview || "",
      body: msg.body?.content || "",
      bodyType: msg.body?.contentType === "html" ? "html" : "text",
      snippet: msg.bodyPreview || "",
      receivedDateTime: msg.receivedDateTime,
      importance: msg.importance?.toLowerCase() || "normal",
      hasAttachments: msg.hasAttachments || false,
      isReminderOn: msg.isReminderOn,
      isFlagged: msg.flag?.flagStatus === "flagged"
    };
  }
}
