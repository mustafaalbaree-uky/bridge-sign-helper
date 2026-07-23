# Automatic email setup

The Review screen can send the notification email automatically. There are two
ways; **Gmail with an app password is the recommended, durable option.**

The actual sending happens in the `notify` Supabase Edge Function, so no email
credentials ever live in this repo or in the browser.

---

## Method A — Gmail app password (recommended)

Uses an existing Gmail account to send over SMTP. App passwords don't expire, so
this keeps working long-term. Best to use an account Google already trusts.

1. **Turn on 2-Step Verification** for the Google account (required for app
   passwords): Google Account → Security → 2-Step Verification.
2. **Create an app password:** Google Account → Security → App passwords →
   create one (name it e.g. "Bridge Sign Helper"). Copy the 16-character code.
3. **Add two secrets in Supabase:** dashboard → your project → **Edge Functions
   → Secrets** (add secret):
   - `GMAIL_USER` = the full sending address (e.g. `you@gmail.com`)
   - `GMAIL_APP_PASSWORD` = the 16-character app password (spaces are fine)
4. In the app: **Setup → Email notifications**, tick **Enable automatic
   sending**, **Save**, then **Send a test email** to confirm.

That's it — the function sends through Gmail whenever those secrets are present.

---

## Method B — Google Apps Script webhook (alternative)

Sends via a script bound to a Google account. Note: brand-new Gmail accounts
used only for automation can get flagged/suspended by Google.

1. Sign in to the sending Gmail. Go to <https://script.google.com> → New project.
2. Paste this, replacing the token with the one shown in **Setup → Email
   notifications → (Apps Script) Shared token**:

   ```javascript
   function doPost(e) {
     var EXPECTED_TOKEN = "PASTE_TOKEN_FROM_SETUP_HERE";
     var data = JSON.parse(e.postData.contents);
     if (data.token !== EXPECTED_TOKEN) {
       return ContentService.createTextOutput(JSON.stringify({ error: "unauthorized" }))
         .setMimeType(ContentService.MimeType.JSON);
     }
     MailApp.sendEmail({ to: data.to, subject: data.subject, body: data.body });
     return ContentService.createTextOutput(JSON.stringify({ ok: true }))
       .setMimeType(ContentService.MimeType.JSON);
   }
   ```

3. **Deploy → New deployment → Web app**, Execute as **Me**, Who has access
   **Anyone** (this exact setting matters — "Anyone with Google Account" causes a
   403). Deploy and authorize.
4. Copy the `/exec` URL. In the app, open the "Alternative: Google Apps Script"
   section, paste the URL + matching token, tick Enable, Save, and test.

If both Gmail secrets and an Apps Script URL are set, Gmail wins.

---

## Turning it off

Untick **Enable automatic sending** and Save. Review falls back to **Compose**
(opens your mail app), which always works with no setup.
