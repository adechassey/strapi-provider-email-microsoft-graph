import { Message } from "@microsoft/microsoft-graph-types";
import { EmailClient } from "./email-client";

interface ProviderOptions {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

interface Settings {
  defaultFrom?: string;
  defaultReplyTo?: string;
  defaultApiSender?: string;
}

interface SendOptions {
  from?: string;
  to: string;
  cc: string;
  bcc: string;
  replyTo?: string;
  subject: string;
  text: string;
  html: string;
  apiSender: string;
  [key: string]: unknown;
}

interface SendBulkOptions {
  from?: string;
  to: string[];
  cc: string;
  bcc: string;
  replyTo?: string;
  subject: string;
  text: string;
  html: string;
  apiSender: string;
  [key: string]: unknown;
}

export default {
  provider: "strapi-provider-email-microsoft-graph",

  init(providerOptions: ProviderOptions, settings: Settings = {}) {
    const emailClient = EmailClient.getInstance(
      providerOptions.tenantId,
      providerOptions.clientId,
      providerOptions.clientSecret
    );

    return {
      async send(options: SendOptions) {
        const apiSender = options.apiSender || settings.defaultApiSender;
        const from = options.from || settings.defaultFrom;
        const replyTo = options.replyTo || settings.defaultReplyTo;

        if (!apiSender) {
          throw new Error("apiSender address is required.");
        }
        if (!from) {
          throw new Error("from address is required.");
        }

        const message: Message = {
          subject: options.subject,
          body: {
            content: options.html || options.text,
            contentType: options.html ? "html" : "text",
          },
          from: {
            emailAddress: {
              address: from,
            },
          },
          replyTo: replyTo
            ? [
                {
                  emailAddress: {
                    address: replyTo,
                  },
                },
              ]
            : undefined,
          toRecipients: [
            {
              emailAddress: {
                address: options.to,
              },
            },
          ],
        };

        await emailClient.sendEmail(message, apiSender);
      },
    };
  },
};
