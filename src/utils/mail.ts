/**
 * Email utility functions for Outlook MCP Server
 */

export function extractEmailAddress(email: string): string {
  const match = email.match(/<(.+?)>/);
  return match ? match[1] : email;
}

export function formatEmailList(recipients: string[]): string {
  return recipients.join(", ");
}

export function truncateText(text: string, maxLength: number = 100): string {
  if (text.length <= maxLength) {
    return text;
  }
  return text.substring(0, maxLength) + "...";
}

export function decodeHtml(html: string): string {
  const entities: Record<string, string> = {
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": '"',
    "&#39;": "'",
    "&#x2F;": "/"
  };

  return html.replace(/&(?:amp|lt|gt|quot|#39|#x2F);/g, (match) => entities[match]);
}

export function isValidMessageId(id: string): boolean {
  return id.length > 0 && typeof id === "string";
}
