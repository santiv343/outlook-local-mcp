# Security

The server uses local stdio and is read-only by default. It provides no remote
authentication, HTTP listener or attachment downloads. Returned mail data reaches
the requesting AI client.

Write and send capabilities require explicit flags at startup, enforced in both
processes. Send previews bind the complete supported content and selected account;
the token is not proof of human consent. Clients must obtain explicit approval.
Concurrent Outlook edits remain best effort, and uncertain mutations must not be
automatically retried. Unsupported HTML, attachments or delegated From identities
require Outlook's own UI. Short references are temporary locators, not access controls.

Report vulnerabilities through GitHub's **Report a vulnerability** option on the
repository's Security tab when available. Use synthetic reproductions and never
publish mailbox data, credentials or personal configuration in issues.

The current release is maintained. Review versioned changes before updating.
Corporate Outlook protections and client permissions remain in effect.
