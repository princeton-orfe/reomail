# reomail

reomail (regarding ORFE mail) is a mailer that accepts a responsive and accessible HTML body and can be used to send styled newsletters, communications, and releases to lists and other groups without the use of a third-party.

The tool is designed to connect directly to a Princeton email account to send the message, much like any other email client application.

## Authentication

AUthentication requires an App registration via Microsoft Azure with the following API permissions.

* Mail.Read
* Mail.Send
* offline_access
* User.Read

The registered App must have a valid Client secret and ID, which you should store alongside your Tenant ID in an untracked file named `.env`.

```
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_id
```
