"""
Applicant Mailer — Streamlit UI

Kept deliberately simple:
  - No SMTP jargon shown to the user (Gmail/Outlook/Yahoo presets handle it
    behind the scenes; "Other" is the only place raw server details show up).
  - No category/status/search filters — just one toggle (New only / Everyone)
    plus a checkbox table people can hand-edit.
  - Honest wording about what "sent" means (see note near the send button).

Run locally:   streamlit run app.py
"""

import json

import streamlit as st
import pandas as pd

import emailer

SENT_LOG_PATH = "sent_log.json"

st.set_page_config(page_title="Applicant Mailer", page_icon="📧", layout="wide")

# --------------------------------------------------------------------------- #
# Session state setup
# --------------------------------------------------------------------------- #

if "recipients_df" not in st.session_state:
    st.session_state.recipients_df = pd.DataFrame(
        columns=["Category", "Name", "Email", "SubmittedAt", "Key", "Send", "PreviouslySent", "LastSent"]
    )
if "parse_report" not in st.session_state:
    st.session_state.parse_report = []
if "sent_log" not in st.session_state:
    st.session_state.sent_log = emailer.load_sent_log(SENT_LOG_PATH)
if "pending_send_resets" not in st.session_state:
    # Keys whose "Send?" checkbox needs to be unchecked, queued up from a
    # previous run (e.g. right after a successful send). We can't write to
    # st.session_state[send_state_key(...)] once that checkbox widget has
    # already been instantiated *in this same run* — Streamlit forbids it.
    # So instead of setting it immediately, we queue the key here and apply
    # the reset on the *next* run, before any checkbox widget is created.
    st.session_state.pending_send_resets = []
if "last_send_results" not in st.session_state:
    st.session_state.last_send_results = None


def send_state_key(key: str) -> str:
    """Session-state key for the individual 'Send' checkbox of one recipient."""
    return f"send::{key}"


def set_send(key: str, value: bool) -> None:
    """
    Programmatically set one recipient's checkbox. Writing directly to
    st.session_state[widget_key] is the reliable way to change a widget's
    value from code in Streamlit — unlike st.data_editor, a plain
    st.checkbox doesn't rely on position-based diffing against a
    re-passed DataFrame, so there's no row-order/key-change edge case
    where the click silently doesn't register.
    """
    st.session_state[send_state_key(key)] = value


def sync_send_column_from_widgets():
    """Pull every recipient checkbox's current widget value back into the
    master recipients_df, so downstream code (counts, the compose tab,
    sending) always sees what's actually checked on screen."""
    df = st.session_state.recipients_df
    if df.empty:
        return
    df["Send"] = df["Key"].apply(lambda k: st.session_state.get(send_state_key(k), False))
    st.session_state.recipients_df = df


def refresh_previously_sent():
    df = st.session_state.recipients_df
    if df.empty:
        return
    log = st.session_state.sent_log
    df["PreviouslySent"] = df["Key"].apply(lambda k: k in log)
    df["LastSent"] = df["Key"].apply(lambda k: log.get(k, {}).get("last_sent", ""))
    st.session_state.recipients_df = df


if st.session_state.pending_send_resets:
    for k in st.session_state.pending_send_resets:
        st.session_state[send_state_key(k)] = False
    st.session_state.pending_send_resets = []

st.title("📧 Applicant Mailer")

tab_upload, tab_recipients, tab_compose, tab_history = st.tabs(
    ["1️⃣ Upload File", "2️⃣ Review Recipients", "3️⃣ Compose & Send", "📜 Sent History"]
)

# --------------------------------------------------------------------------- #
# TAB 1 — Upload
# --------------------------------------------------------------------------- #
with tab_upload:
    st.subheader("Upload your Excel file")
    uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])

    if uploaded is not None:
        # st.file_uploader keeps returning the same file on EVERY rerun of
        # the app, not just the run right after it was picked — including
        # reruns triggered by something unrelated, like ticking a checkbox
        # over in the Review Recipients tab. Without this guard, this whole
        # block (which resets every "Send" checkbox to its default) would
        # re-run on every single click anywhere in the app, instantly
        # undoing whatever the person had just checked/unchecked. Only
        # (re)process when it's actually a new file.
        upload_signature = getattr(uploaded, "file_id", None) or (uploaded.name, uploaded.size)
        if st.session_state.get("last_upload_signature") != upload_signature:
            st.session_state.last_upload_signature = upload_signature
            file_bytes = uploaded.getvalue()
            new_df, report = emailer.load_recipients_from_workbook(file_bytes)
            st.session_state.parse_report = report

            if new_df.empty:
                st.error("No email addresses were found in this file.")
            else:
                old_df = st.session_state.recipients_df
                if not old_df.empty:
                    combined = pd.concat([old_df[["Category", "Name", "Email", "SubmittedAt", "Key"]], new_df])
                    combined = combined.drop_duplicates(subset="Key", keep="last").reset_index(drop=True)
                else:
                    combined = new_df

                st.session_state.recipients_df = combined
                refresh_previously_sent()
                # New people default to checked; anyone already emailed defaults to unchecked.
                for _, row in st.session_state.recipients_df.iterrows():
                    set_send(row["Key"], not row["PreviouslySent"])
                sync_send_column_from_widgets()

                st.success(f"Loaded {len(new_df)} recipient(s). Total loaded: {len(st.session_state.recipients_df)}.")

        if st.session_state.parse_report:
            with st.expander("Details (which columns were used per sheet)", expanded=False):
                for r in st.session_state.parse_report:
                    if r.used:
                        st.write(f"✅ **{r.sheet_name}** — {r.reason} (email: *{r.email_header}*, name: *{r.name_header}*)")
                    else:
                        st.write(f"⏭️ **{r.sheet_name}** — skipped ({r.reason})")

    with st.expander("Restore a previous sent-history file", expanded=False):
        restore_log = st.file_uploader("sent_log.json", type=["json"], key="restore_log_uploader")
        if restore_log is not None:
            restore_signature = getattr(restore_log, "file_id", None) or (restore_log.name, restore_log.size)
            if st.session_state.get("last_restore_signature") != restore_signature:
                st.session_state.last_restore_signature = restore_signature
                try:
                    restored = json.loads(restore_log.getvalue())
                    st.session_state.sent_log.update(restored)
                    emailer.save_sent_log(SENT_LOG_PATH, st.session_state.sent_log)
                    refresh_previously_sent()
                    st.success(f"Restored {len(restored)} entries.")
                except Exception as e:
                    st.error(f"Couldn't read that file: {e}")

# --------------------------------------------------------------------------- #
# TAB 2 — Review recipients
# --------------------------------------------------------------------------- #
with tab_recipients:
    df = st.session_state.recipients_df

    if df.empty:
        st.info("Upload a file in the **Upload File** tab first.")
    else:
        st.subheader("Choose who to send to")

        show_choice = st.radio(
            "Show",
            ["New only (not yet emailed)", "Everyone"],
            index=1,
            horizontal=True,
            label_visibility="collapsed",
        )
        view = df if show_choice == "Everyone" else df[~df["PreviouslySent"]]

        # Make sure every visible recipient has a checkbox widget value to
        # start from (first time we see them this session).
        for k in view["Key"]:
            state_key = send_state_key(k)
            if state_key not in st.session_state:
                set_send(k, bool(df.loc[df["Key"] == k, "Send"].iloc[0]))

        btn1, btn2, _ = st.columns([1, 1, 4])
        if btn1.button("☑️ Select all"):
            for k in view["Key"]:
                set_send(k, True)
            st.rerun()
        if btn2.button("⬜ Deselect all"):
            for k in view["Key"]:
                set_send(k, False)
            st.rerun()

        if view.empty:
            st.info("Nobody matches this view. Everyone has already been emailed — switch to **Everyone** to resend.")
        else:
            header = st.columns([0.6, 1.6, 1.8, 2.6, 1.6])
            for col, label in zip(header, ["Send?", "Sheet", "Name", "Email", "Already emailed?"]):
                col.markdown(f"**{label}**")
            for _, row in view.iterrows():
                c1, c2, c3, c4, c5 = st.columns([0.6, 1.6, 1.8, 2.6, 1.6])
                c1.checkbox(
                    "Send?",
                    key=send_state_key(row["Key"]),
                    label_visibility="collapsed",
                )
                c2.write(row["Category"])
                c3.write(row["Name"] or "—")
                c4.write(row["Email"])
                c5.write("✅" if row["PreviouslySent"] else "—")

        sync_send_column_from_widgets()
        df = st.session_state.recipients_df

        st.caption(f"**{int(df['Send'].sum())}** of {len(df)} recipients selected.")

        st.download_button(
            "⬇️ Download recipient list as CSV",
            df.drop(columns=["Key"]).to_csv(index=False).encode("utf-8"),
            file_name="recipients.csv",
            mime="text/csv",
        )

# --------------------------------------------------------------------------- #
# TAB 3 — Compose & send
# --------------------------------------------------------------------------- #
with tab_compose:
    df = st.session_state.recipients_df
    selected = df[df["Send"] == True] if not df.empty else df

    st.subheader("Send from")
    c1, c2 = st.columns(2)
    with c1:
        provider = st.selectbox("Email provider", list(emailer.SMTP_PRESETS.keys()))
        sender_email = st.text_input("Your email address", placeholder="you@gmail.com")
        sender_name = st.text_input("Display name (optional)", placeholder="YTUMUN Team")
    with c2:
        sender_password = st.text_input(
            "App password",
            type="password",
            help="Not your normal password — a generated app password. Nothing typed here is saved.",
        )
        if provider == "Other":
            smtp_host = st.text_input("Server address", placeholder="smtp.example.com")
            smtp_port = st.number_input("Port", value=587)
            security = st.selectbox("Security", ["STARTTLS", "SSL", "None"])
        else:
            preset = emailer.SMTP_PRESETS[provider]
            smtp_host, smtp_port, security = preset["host"], preset["port"], preset["security"]

    st.divider()
    st.subheader("Email content")
    st.caption("Use **{name}**, **{first_name}**, or **{email}** anywhere — they'll be filled in per person.")

    subject_template = st.text_input("Subject", value="YTUMUN: Constellation Edition — Application Received")
    body_template = st.text_area(
        "Body",
        height=280,
        value=(
            "Dear {first_name},\n\n"
            "Thank you for applying to YTUMUN: Constellation Edition! We've received your "
            "application and our team will be in touch soon.\n\n"
            "Best regards,\n"
            "YTUMUN Team"
        ),
    )

    if selected is not None and len(selected) > 0:
        sample = selected.iloc[0]
        with st.expander("👀 Preview (first selected recipient)"):
            st.write(f"**To:** {sample['Email']}")
            st.write(f"**Subject:** {emailer.personalize(subject_template, sample['Name'], sample['Email'])}")
            st.text(emailer.personalize(body_template, sample['Name'], sample['Email']))

    st.divider()
    colA, colB = st.columns(2)

    with colA:
        if st.button("✉️ Send test email to myself"):
            if not sender_email or not sender_password:
                st.error("Fill in your sender email and app password first.")
            else:
                test_name = selected.iloc[0]["Name"] if selected is not None and len(selected) > 0 else "Test Person"
                try:
                    with emailer.SMTPSession(smtp_host, int(smtp_port), security, sender_email, sender_password) as session:
                        result = session.send(
                            sender_email, sender_name, sender_email,
                            "[TEST] " + emailer.personalize(subject_template, test_name, sender_email),
                            emailer.personalize(body_template, test_name, sender_email),
                        )
                    if result.success:
                        st.success(f"Test email sent to {sender_email}. Check your inbox.")
                    else:
                        st.error(f"Failed: {result.error}")
                except Exception as e:
                    st.error(f"Couldn't connect/authenticate: {e}")

    with colB:
        n_selected = int(selected["Send"].sum()) if selected is not None and not selected.empty else 0
        send_clicked = st.button(f"🚀 Send to {n_selected} selected recipient(s)", type="primary")

    st.caption(
        "ℹ️ A green checkmark below means the mail server **accepted** the message — not a delivery "
        "guarantee. If an address is fake or doesn't exist, most providers (including Gmail) bounce it "
        "back later as a separate email to your own inbox, sometimes several minutes afterward — not "
        "as an in-app failure. Check your inbox after sending a batch."
    )

    if send_clicked:
        if df.empty or selected.empty:
            st.error("No recipients selected. Go to the Review Recipients tab and check some boxes.")
        elif not sender_email or not sender_password:
            st.error("Fill in your sender email and app password first.")
        elif not subject_template.strip() or not body_template.strip():
            st.error("Subject and body can't be empty.")
        else:
            recipients_payload = [
                {
                    "key": row["Key"],
                    "name": row["Name"],
                    "email": row["Email"],
                    "subject_template": subject_template,
                    "body_template": body_template,
                }
                for _, row in selected.iterrows()
            ]

            progress_bar = st.progress(0, text="Starting...")
            status_area = st.empty()
            log_lines = []

            def on_progress(i, total, result):
                progress_bar.progress(i / total, text=f"Processed {i}/{total}")
                mark = "✅ accepted" if result.success else "❌ rejected"
                log_lines.append(f"{mark} — {result.email}" + (f" ({result.error})" if result.error else ""))
                status_area.text("\n".join(log_lines[-15:]))

            try:
                results = emailer.send_batch(
                    smtp_host, int(smtp_port), security,
                    sender_email, sender_password, sender_name, "",
                    recipients_payload, False, 1.0,
                    progress_callback=on_progress,
                )
            except Exception as e:
                st.error(f"Couldn't connect/authenticate: {e}")
                results = []

            if results:
                for payload, result in zip(recipients_payload, results):
                    if result.success:
                        subj = emailer.personalize(payload["subject_template"], payload["name"], payload["email"])
                        emailer.mark_sent(st.session_state.sent_log, payload["key"], subj)
                emailer.save_sent_log(SENT_LOG_PATH, st.session_state.sent_log)
                refresh_previously_sent()

                # Don't touch st.session_state[send_state_key(...)] here — the
                # "Send?" checkboxes for these same keys were already
                # instantiated earlier in this run (in the Review Recipients
                # tab), and Streamlit raises a StreamlitAPIException if you
                # write to a widget's state after it's been created in the
                # same run. Instead, queue the keys and let the reset happen
                # at the top of the *next* run, before those widgets exist.
                st.session_state.pending_send_resets = [
                    payload["key"] for payload, result in zip(recipients_payload, results) if result.success
                ]

                results_df = pd.DataFrame(
                    [{"Email": r.email, "Accepted": r.success, "Error": r.error} for r in results]
                )
                st.session_state.last_send_results = {
                    "n_ok": sum(1 for r in results if r.success),
                    "n_fail": sum(1 for r in results if not r.success),
                    "results_df": results_df,
                }
                st.rerun()

    if st.session_state.last_send_results:
        info = st.session_state.last_send_results
        st.success(f"Done. {info['n_ok']} accepted by the mail server, {info['n_fail']} rejected immediately.")
        st.dataframe(info["results_df"], hide_index=True, use_container_width=True)
        st.download_button(
            "⬇️ Download send results as CSV",
            info["results_df"].to_csv(index=False).encode("utf-8"),
            file_name="send_results.csv",
            mime="text/csv",
        )

# --------------------------------------------------------------------------- #
# TAB 4 — Sent history
# --------------------------------------------------------------------------- #
with tab_history:
    st.subheader("Sent history")
    st.caption("Who's already been emailed. Download this after sending — see below for why.")

    log = st.session_state.sent_log
    if not log:
        st.info("No emails sent yet.")
    else:
        rows = []
        for key, entry in log.items():
            category, email = (key.split("||", 1) + [""])[:2]
            rows.append(
                {
                    "Sheet": category,
                    "Email": email,
                    "Times sent": entry.get("times_sent", 0),
                    "Last sent": entry.get("last_sent", ""),
                }
            )
        hist_df = pd.DataFrame(rows).sort_values("Last sent", ascending=False)
        st.dataframe(hist_df, hide_index=True, use_container_width=True)

        st.download_button(
            "⬇️ Download sent_log.json (back this up!)",
            json.dumps(log, indent=2, default=str).encode("utf-8"),
            file_name="sent_log.json",
            mime="application/json",
        )

    st.caption(
        "⚠️ On free hosting this file can be wiped when the app restarts. Download it after each "
        "batch and re-upload it (Upload File tab) if the app ever forgets."
    )
