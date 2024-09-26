# Strapi Microsoft Graph Email Provider

This project provides a custom email provider for Strapi using Microsoft Graph API. It allows sending emails through Microsoft Graph, leveraging the powerful features and integration capabilities of Azure services.

## Features

- Send emails using Microsoft Graph API
- Support for custom `from`, `replyTo`, and `apiSender` addresses
- Support for HTML and text email bodies
- Ability to add attachments
- Easy configuration and initialization

## Installation

To install the Microsoft Graph Email Provider, run the following command:

```bash
npm install strapi-provider-email-microsoft-graph
```

## Configuration

You need to configure the provider with your Microsoft Graph API credentials. Add the following configuration to your Strapi project.

### Provider Options

| Option         | Type   | Description                    |
| -------------- | ------ | ------------------------------ |
| `tenantId`     | string | Your Azure AD tenant ID        |
| `clientId`     | string | Your application client ID     |
| `clientSecret` | string | Your application client secret |

### Settings

| Option             | Type   | Description                      |
| ------------------ | ------ | -------------------------------- |
| `defaultFrom`      | string | Default `from` email address     |
| `defaultReplyTo`   | string | Default `replyTo` email address  |
| `defaultApiSender` | string | Default API sender email address |

### Example Configuration

```js
// config/plugins.js

module.exports = ({ env }) => ({
  email: {
    config: {
      provider: "strapi-provider-email-microsoft-graph",
      providerOptions: {
        tenantId: env("MICROSOFT_GRAPH_TENANT_ID"),
        clientId: env("MICROSOFT_GRAPH_CLIENT_ID"),
        clientSecret: env("MICROSOFT_GRAPH_CLIENT_SECRET"),
      },
      settings: {
        defaultFrom: "no-reply@example.com",
        defaultReplyTo: "support@example.com",
        defaultApiSender: "api@example.com",
      },
    },
  },
});
```

## Usage

### Sending an Email

To send an email, use the `send` function provided by the email provider. Here is an example:

```js
// controllers/email.js

module.exports = {
  async sendEmail(ctx) {
    const { to, subject, text, html, attachments } = ctx.request.body;

    try {
      await strapi.plugins["email"].services.email.send({
        to,
        subject,
        text,
        html,
        attachments,
      });
      ctx.send({ message: "Email sent successfully" });
    } catch (err) {
      ctx.send({ error: "Failed to send email" });
    }
  },
};
```

### Send Options

| Option        | Type                                                                    | Description                          |
| ------------- | ----------------------------------------------------------------------- | ------------------------------------ |
| `to`          | string                                                                  | Recipient email address              |
| `subject`     | string                                                                  | Subject of the email                 |
| `from`        | string (optional if `defaultFrom` is defined in provider settings)      | Email address of the sender          |
| `apiSender`   | string (optional if `defaultApiSender` is defined in provider settings) | API sender email address             |
| `cc`          | string (optional)                                                       | CC email addresses, comma-separated  |
| `bcc`         | string (optional)                                                       | BCC email addresses, comma-separated |
| `replyTo`     | string (optional)                                                       | Reply-to email address               |
| `text`        | string (optional)                                                       | Text body of the email               |
| `html`        | string (optional)                                                       | HTML body of the email               |
| `attachments` | array (optional)                                                        | Attachments to include in the email  |

## Contributing

Contributions are welcome! If you have any improvements or suggestions, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---
