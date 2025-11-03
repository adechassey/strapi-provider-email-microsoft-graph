import { Message, Recipient } from "@microsoft/microsoft-graph-types";
import { EmailClient } from "./email-client";

interface ProviderOptions {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

interface Settings {
  defaultFrom?: string;
  defaultReplyTo?: string | string[];
  defaultApiSender?: string;
}

interface SendOptions {
  from?: string;
  to: string | string[];
  cc: string;
  bcc: string;
  replyTo?: string | string[];
  subject: string;
  text: string;
  html: string;
  apiSender: string;
  attachments?: Message["attachments"];
}

export default {
  provider: "strapi-provider-email-microsoft-graph",
  name: "Microsoft Graph Email Provider",

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

        const toRecipients: Recipient[] = [options.to]
          .flat()
          .map((address) => ({
            emailAddress: { address },
          }));

        const replyToRecipients: Recipient[] = [replyTo]
          .flat()
          .map((address) => ({
            emailAddress: { address },
          }));

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
          replyTo: replyToRecipients
            ? replyToRecipients
            : undefined,
          toRecipients,
          attachments: options.attachments,
        };

        await emailClient.sendEmail(message, apiSender);
      },
    };
  },
};
