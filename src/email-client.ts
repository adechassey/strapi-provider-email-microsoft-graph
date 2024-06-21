import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Message } from "@microsoft/microsoft-graph-types";

export class EmailClient {
  private static instance: EmailClient;
  private client: Client;

  constructor(tenantId: string, clientId: string, clientSecret: string) {
    const clientSecretCredential = new ClientSecretCredential(
      tenantId,
      clientId,
      clientSecret
    );

    const authProvider = new TokenCredentialAuthenticationProvider(
      clientSecretCredential,
      { scopes: ["https://graph.microsoft.com/.default"] }
    );

    this.client = Client.initWithMiddleware({ authProvider });
  }

  static getInstance(tenantId: string, clientId: string, clientSecret: string) {
    if (!EmailClient.instance) {
      EmailClient.instance = new EmailClient(tenantId, clientId, clientSecret);
    }

    return EmailClient.instance;
  }

  async sendEmail(message: Message, apiSender: string) {
    try {
      await this.client.api(`/users/${apiSender}/sendMail`).post({ message });
    } catch (error) {
      throw error;
    }
  }
}
