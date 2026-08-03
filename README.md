# Applicant Mailer

A small Streamlit app for sending personalized "thanks for applying" (or any
other) emails to everyone who filled out your Google Form, using the Excel
sheet the form writes to.

**What it does**
- Reads your `.xlsx` response sheet. It scans **every tab**, and on each tab
  automatically figures out which column is the email address and which is
  the applicant's name — it does not assume fixed column positions, so it
  keeps working even if you add/remove/reorder form questions later.
- Lets you write **one** email with `{name}`, `{first_name}`, and `{email}`
  placeholders, which get swapped in per person.
- Lets you type in whichever email address you want to send *from* (your own
  for testing; a friend can type in theirs when you hand the app to them —
  nothing is hard-coded).
- Remembers who's already been emailed, so if you upload an updated sheet
  later (with new applicants added), it automatically pre-checks only the
  new people. You can always override the checkboxes by hand.

---

## 1. Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Your browser will open to a local page with four tabs: **Upload File →
Review Recipients → Compose & Send → Sent History**. Work through them top
to bottom.

## 2. Set up an "app password" (Gmail / Outlook / Yahoo)

Most providers no longer accept your normal account password from a script
— you need a separate **app password**. This is a one-time setup per
sender email address.

**Gmail**
1. Turn on 2-Step Verification: `myaccount.google.com/security`
2. Go to `myaccount.google.com/apppasswords`
3. Create an app password (name it anything, e.g. "Applicant Mailer")
4. Copy the 16-character password and paste it into the app's "App
   password" field (not your real Gmail password)

**Outlook / Office365**
1. Turn on 2-Step Verification at `account.microsoft.com/security`
2. Create an "app password" under Security → Advanced security options

**Yahoo**
1. Account Security → Generate app password

If your friend uses a different provider, pick **Other** in the app and
enter their server address/port manually (their IT/email provider's help
page will list it).

## 3. Deploy it so your friend can use it

The easiest free option is **Streamlit Community Cloud**:

1. Push this folder (`app.py`, `emailer.py`, `requirements.txt`) to a
   GitHub repo (it can be private).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click "New app", pick the repo and `app.py` as the entry point.
3. Deploy. You'll get a public URL you can send to your friend.

Nobody's email password is stored anywhere in the code or repo — each
person types their own app password into the running app, and it only
lives in that browser session's memory.

### ⚠️ Important limitation: the "already emailed" list

The app tracks who's been emailed in a local file (`sent_log.json`) next to
the app. On Streamlit Community Cloud's free tier, this file can be wiped
whenever the app restarts or you push a new version. To be safe:

- After sending a batch, go to the **Sent History** tab and click
  **Download sent_log.json** to back it up.
- If the app ever "forgets" (e.g. after a redeploy), upload that file back
  in the **Upload File** tab under "Restore a previous sent-history file" —
  it'll re-mark everyone in it as already emailed.

If you outgrow this, the natural next step is swapping the JSON file for a
small database (e.g. a free Google Sheet or Supabase table) — the
`load_sent_log` / `save_sent_log` / `mark_sent` functions in `emailer.py`
are the only three functions that would need to change.

## 4. What a green "sent" checkmark actually means

When the app shows a recipient as accepted, that means the mail server took
the message — it is **not** a delivery guarantee. If an address is fake or
doesn't exist, most providers (Gmail included) accept it first and only
bounce it back later, as a separate email sent to *your own* inbox — often
several minutes afterward. This is normal behavior across virtually all
email systems, not a bug in the app. Check your inbox for bounce
notifications after sending a batch.

## 5. Sending limits

Regular Gmail/Outlook/Yahoo accounts typically cap you around 300–500
emails per day. The app adds a small delay between sends (adjustable in the
UI) to stay well under provider rate limits. For a large MUN conference
(hundreds of applicants), consider spreading sends across a couple of days,
or using a transactional email service (SendGrid, Mailgun, etc.) if you
outgrow a normal inbox — that would be a small change to the SMTP settings
in the "Custom" provider option.

## Files

- `app.py` — the Streamlit UI
- `emailer.py` — all the logic (column detection, personalization, sending,
  sent-log tracking) — kept separate so it can be tested independently
- `requirements.txt` — Python dependencies
- `sent_log.json` — created automatically the first time you send an email;
  this is your "already emailed" history
