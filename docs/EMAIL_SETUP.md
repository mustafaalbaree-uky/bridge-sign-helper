# Automatic email setup (Gmail + Google Apps Script)

This lets the **Review** screen send the notification email for real, instead of
just opening your mail app. The email is sent by a Google Apps Script that runs as
a Gmail account you own, so your Gmail password never leaves Google. The app talks
to it through a small server function, so there are no browser security issues.

## How it flows

```
Review screen  ->  notify (Supabase function)  ->  your Apps Script  ->  Gmail sends
```

## One-time setup (about 5 minutes)

1. **Make/choose the sending Gmail.** A dedicated account is fine (e.g.
   `bridgesignbot@gmail.com`). Sign in to it.

2. Go to <https://script.google.com> and click **New project**.

3. Delete the sample code and paste this in, then **replace the token** with the
   value shown in the app under **Setup → Email notifications → Shared token**:

   ```javascript
   function doPost(e) {
     var EXPECTED_TOKEN = "PASTE_TOKEN_FROM_SETUP_HERE";
     var data = JSON.parse(e.postData.contents);
     if (data.token !== EXPECTED_TOKEN) {
       return ContentService
         .createTextOutput(JSON.stringify({ error: "unauthorized" }))
         .setMimeType(ContentService.MimeType.JSON);
     }
     MailApp.sendEmail({ to: data.to, subject: data.subject, body: data.body });
     return ContentService
       .createTextOutput(JSON.stringify({ ok: true }))
       .setMimeType(ContentService.MimeType.JSON);
   }
   ```

4. **Deploy it:** click **Deploy → New deployment**. For type, pick **Web app**.
   - **Execute as:** Me (the Gmail account)
   - **Who has access:** Anyone
   - Click **Deploy**, then **Authorize access** and allow the permissions.

   "Anyone" is required so the server can reach it; the token is what actually
   protects it.

5. **Copy the Web app URL** (it ends in `/exec`).

6. Back in the app: **Setup → Email notifications**. Paste the **URL**, make sure
   the **token** matches what you put in the script, click **Save email settings**,
   then **Send a test email** to yourself to confirm it works.

## Changing or revoking it

- To rotate the token: change it in both the script (redeploy) and the app.
- To turn it off: clear the URL in Setup and save; Review falls back to composing
  in your mail app.
