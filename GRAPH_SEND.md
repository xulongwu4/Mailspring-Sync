# Office 365 send via Microsoft Graph (local patch)

Branch: `graph-send` (based on mailsync `675e9f22`, the version pinned by Mailspring 1.26.0).

## Why

Mailspring signs in to Office 365 with OAuth, reads mail over IMAP and sends over SMTP.
Many tenants turn off SMTP AUTH (Security Defaults, or `SmtpClientAuthenticationDisabled`), so adding the
account fails with:

    Authentication Error - Check your username and password. (SMTP)

IMAP still works. This patch sends through Microsoft Graph `POST /v1.0/me/sendMail` instead,
so no SMTP is used for these accounts. Reading still uses IMAP.

Upstream discussion: community.getmailspring.com/t/institutional-office365-not-working/48 (Graph
support is planned upstream but not implemented), /t/make-smtp-or-imap-optional-read-only-or-send-only-accounts/1.

## How it works

Only mailsync (the C++ sync engine) changes. The Electron app is untouched.

- **Token**: the account's existing refresh token (Mailspring client id `8787a430-6eee-41e1-b914-681d90d35625`)
  is exchanged for a second access token with scope `https://graph.microsoft.com/Mail.Send`.
  Mailspring's onboarding already asks for Graph scopes, and the user must have consented to `Mail.Send`
  once (see "Consent test" below).
- **Which accounts**: `SendsViaGraph(account)` = provider is `office365` **and** the SMTP username is empty or
  equal to the IMAP username. Shared mailboxes (SMTP user = mailbox owner) keep SMTP, because
  `/me/sendMail` can only send as the signed-in user.
- **Send**: the MIME message mailsync already builds is base64-encoded and posted with
  `Content-Type: text/plain`. Graph takes the recipients (To/Cc/Bcc) from the MIME headers and
  returns 202. Graph saves the copy in Sent Items; the existing Sent-folder scan by Message-ID
  finds it like it finds a server-filed SMTP copy.
- **Errors**: any failure (token refresh, HTTP error, network) is raised as
  `SyncException("send-failed", "Microsoft Graph sendMail: ...")`, which the app already shows as
  a send failure.
- **Account check** (`--mode test` during onboarding): for Graph accounts the SMTP login is replaced by
  fetching a `Mail.Send` token. Failure is reported as an SMTP authentication error, with the Graph
  error in the log.

## Changed files

| File | Change |
|------|--------|
| `MailSync/NetworkRequestUtils.hpp/.cpp` | `MakeOAuthRefreshRequest(..., scope = "")` (Microsoft: empty = IMAP+SMTP scopes as before); new `SendsViaGraph()`, `GraphSendMIME()` |
| `MailSync/XOAuth2TokenManager.hpp/.cpp` | `partsForAccount(account, scope = "")`; cache key is `account id + scope` |
| `MailSync/TaskProcessor.cpp` | `performRemoteSendDraft` and the RSVP send use Graph when `SendsViaGraph` |
| `MailSync/main.cpp` | `runTestAuth`: Graph token check instead of SMTP for Graph accounts |

No new source files, so CMakeLists.txt / Xcode / Visual Studio projects need no change.

## Limits

- **4 MB request cap**: Graph rejects larger requests (413). With base64 overhead that is about
  3 MB of MIME, so big attachments fail. Fix: create a draft, upload attachments with an upload
  session, then send the draft.
- **No per-recipient bodies**: when the app asks for a multisend (link/open tracking), everyone gets
  the "self" body, untracked; a warning is logged.
- **Shared mailboxes / send-as** still use SMTP and will still fail if SMTP AUTH is off.
  Fix: `POST /users/{mailbox}/sendMail` with `Mail.Send.Shared` consent.
- **Personal Outlook.com accounts** (provider `outlook`) are not switched; they keep SMTP.
- **Linux only tested**. Code is portable but only the Linux binary was built and tried.
- **No automated test**: every path needs a live Microsoft account. Tested by hand: add account,
  send mail, mail arrives, one copy in Sent.
- **Overwritten on upgrade**: a new Mailspring release ships its own `mailsync.bin`. Rebase this branch onto
  the commit the new release pins, rebuild, and re-deploy.

## Consent test

`graph_consent_test.py` (next to this file, Python stdlib only) runs the same OAuth flow as Mailspring,
prints the granted scopes and checks for `Mail.Send`. With `--send` it sends a test mail via Graph.
If `Mail.Send` is missing, a tenant admin must grant consent to the Mailspring app.

## Build (Linux)

The host (Solus) lacks several build dependencies, so build in Docker like upstream CI (ubuntu 22.04):

    mkdir -p ~/src/mailsync-build
    docker run --rm -v ~/src/Mailspring-Sync:/src:ro -v ~/src/mailsync-build:/out \
      -v ~/src/Mailspring-Sync/build-in-docker.sh:/b.sh:ro ubuntu:22.04 /b.sh

Output: `~/src/mailsync-build/mailsync.bin` (about 1.5 min). Most libraries are linked statically;
it needs libssl3, libcurl, glib, libsasl2 (bundled with Mailspring) and, at runtime, libtidy
(`sudo eopkg it tidy-html5` on Solus).

## Deploy / rollback

Mailspring runs `<resources>/app.asar.unpacked/mailsync` (a wrapper script), which runs `mailsync.bin` in the
same folder (`app/src/mailsync-process.ts` in the Mailspring repo).

    D=~/Downloads/mailspring/usr/share/mailspring/resources/app.asar.unpacked
    cp -n $D/mailsync.bin $D/mailsync.bin.orig      # backup once
    cp ~/src/mailsync-build/mailsync.bin $D/mailsync.bin

Rollback: `cp $D/mailsync.bin.orig $D/mailsync.bin`. Restart Mailspring after either.

Check it is active: after a send the mailsync log shows
`-- Sending a single message body to all recipients via Microsoft Graph:`.
