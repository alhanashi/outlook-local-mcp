# Outlook add-in: mail assistant

A task pane for Outlook (Legacy Outlook for Mac, Outlook for Windows, Outlook on the web) that works with **on-premises Exchange**. It reads the **open message only**, then:

- summarizes it,
- classifies it (needs action / FYI / vendor / ignore) with urgency,
- proposes replies, or writes one from your instruction,
- opens the reply in Outlook's own reply window. **It never sends mail.**

## How it works

- Permission: `ReadItem` only (no mailbox-wide access, no Microsoft Graph).
- The pane calls the Claude API directly from the browser with **your own API key**.
- The key is stored in the add-in's local storage on that device and is sent only to `api.anthropic.com`.
- Email content is treated as untrusted; the prompts tell the model not to follow instructions inside emails.

## Host it

The pane is static files. Host this `addin/` folder on any HTTPS origin (GitHub Pages works), then build the manifest for that URL:

```sh
sh build-manifest.sh https://<user>.github.io/<repo>/addin
```

## Install (per user)

Outlook on the web → Settings → **Manage add-ins** → **+** → **Add from a file** → choose `manifest.xml`. The button **مساعد البريد** then appears when reading a message (also in desktop Outlook after sync).

## First use

Open the pane → **الإعدادات** → paste an API key from the Claude Console → **حفظ**. Pick a model.

## Security notes

- Whoever controls the hosting origin controls the code that runs in the pane. Protect that account (2FA) or host it internally.
- Mail content is sent to the Claude API. Check your organization's data policy before use.
