# Quote Required – Outlook add-in

A one-click Outlook button for Teltherm Instruments and Homershams staff. It marks the open email as needing a quote, which logs it to the **Quote Requests** SharePoint list (and later to HubSpot).

Hosted on GitHub Pages: `https://jason4inz.github.io/quote-addin/`

---

## How it works

```
Rep clicks "Quote Required"  ──►  Add-in (Graph, as the signed-in user)
                                    • adds "Quote Required" category
                                    • flags the email
                                          │
                                          ▼
                                 Power Automate: flw_QuoteReq_Flag
                                 (one copy per mailbox)
                                    • creates item in Quote Requests list
                                    • completes the flag
                                    • swaps category to "Quote Logged"
                                    • moves the email to the logged folder
```

| Piece | Where |
|---|---|
| Add-in files | This repo (`manifest.xml`, `commands.html`, `commands.js`, `assets/`) |
| Hosting | GitHub Pages, public repo `jason4INZ/quote-addin` |
| Authentication | Entra app registration **"Teltherm Quote Required add-in"** (admin consent granted). Client and tenant IDs are in `commands.js` |
| Graph scopes | `Mail.ReadWrite`, `MailboxSettings.ReadWrite` (delegated) |
| Deployment | M365 admin centre → Settings → **Integrated apps** |
| Flow | `flw_QuoteReq_Flag` in Power Automate, which fires when an email is flagged |
| Destination list | **Quote Requests** SharePoint list (Teltherm site) |

The add-in code works on whoever is signed in (`/me`), so it needs no per-user changes. The only per-user part is the flow.

---

## Adding a new user

### 1. Assign the add-in
1. Go to admin.microsoft.com → **Settings → Integrated apps**.
2. Open **Quote Required**, then **Users** → add the person → save.
3. The button shows up under a **Teltherm** group on the ribbon of an open email. It can take up to 24 hours to appear.

You don't need to change the app registration, because admin consent already covers the whole organisation.

### 2. Create their copy of the flow
The Outlook trigger only watches the mailbox of whoever owns the connection, so each user needs their own copy.

1. In Power Automate, open **flw_QuoteReq_Flag** → **Save As**.
2. Name the copy after the user, e.g. `flw_QuoteReq_Flag_<Initials>`.
3. In the copy, change the **Office 365 Outlook** connection to one **signed in as the new user**. The trigger, the category change and the email move all need to run as them.
4. You can leave the **SharePoint** connection as the existing one, so the user doesn't need edit rights on the list.
5. Turn the flow **on**.

### 3. Prepare their mailbox
- Create the logged folder the flow moves emails into: **`Quote Logged`**. The name must match the flow exactly.
- Optional: add a **Quote Logged** category so it has a colour. The add-in creates the **Quote Required** category automatically on first use.
- Optional: set up the **Ctrl+Shift+5** Quick Step fallback (applies the "Quote Required" category and flags the email). It's used on Outlook versions that can't run the add-in. Quick Steps belong to each mailbox, so set it up in their Outlook too.

### 4. Test
1. In the user's Outlook, open a test email and click **Quote Required**.
2. A "Marked Quote Required" message should appear on the email.
3. Check that a new item appears in the Quote Requests list.
4. Check that the email has moved to the logged folder with the **Quote Logged** category.

Keep the list of who has been set up in the internal Teltherm records, not in this public repo.

---

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| Button not showing | Add-in not yet assigned, or still propagating (up to 24 h). Restart Outlook. |
| "This version of Outlook can't run the Quote Required button" | Outlook doesn't support Nested App Authentication. Update Outlook or use the Ctrl+Shift+5 Quick Step. |
| "Quote Required failed: Graph 401/403" | Consent or sign-in problem. Check that the app registration still has admin consent for both scopes. |
| Email flagged but nothing logged | The user's flow copy is missing or off, or its Outlook connection isn't theirs. |
| Flow fails at the move step | The logged folder doesn't exist in their mailbox, or its name doesn't match. |

---

## Updating the add-in
1. Edit the files in this repo and commit to `main`. GitHub Pages republishes within a minute or two.
2. If you change `manifest.xml` (for example new buttons or icons), bump `<Version>`. Then update the app in **Integrated apps** so the change reaches every user.
3. The category names in `commands.js` (`Quote Required`, `Quote Logged`) must match the flow conditions exactly.

## Roadmap
- AI Builder prompt to extract details from the email, matching the sender to the AZUL customer account.
- Push records from the Quote Requests list into HubSpot.
