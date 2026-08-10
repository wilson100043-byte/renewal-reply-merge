# Security notes

- Never commit Google OAuth access tokens, refresh tokens, Client Secrets, customer data, or production Excel files.
- `config.js` may contain a Web OAuth Client ID. A Client ID is public metadata; it must still be restricted to the intended JavaScript origins in Google Cloud Console.
- The application requests Google Drive read-only access and does not write back to Drive.
- Excel contents are processed in browser memory and are only written to a user-initiated download.
- GitHub Pages is publicly reachable by default. Do not treat repository privacy as website access control.
