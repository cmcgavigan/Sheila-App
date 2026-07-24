# Deploying Sheila-App for your wife (remote, phone-only)

Her phone is the client; **your laptop is the server**. Receipts she takes while
your laptop is off/asleep queue on her phone and file automatically when it's back.

Her app address (Tailscale, on its own HTTPS port so it doesn't clash with your
Receipts app on :443):

    https://desktop-rh302uu.tail999c81.ts.net:8443

---

## Part A — things ONLY YOU can do (Tailscale account actions)

These are account/admin actions on **your** Tailscale. They can't be automated.

**Join method = invite + "Sign in with Apple" (easiest on iPhone).**
NOTE: Tailscale's iOS app has NO QR/deep-link for auth keys, so the auth-key
route would mean her pasting a long key by hand. The invite route is smoother:
she signs in with her own Apple ID (one Face ID tap) and accepts your invite.

1. **Invite her to your tailnet.** In the admin console (login.tailscale.com/admin):
   - **Users → Invite external users → enter her email**, OR **Machines → Share**
     a device with her. She gets an email invite.
   - She installs Tailscale, taps **Sign in with Apple**, then opens that email
     and taps **Accept**. Her phone is now on your tailnet.
   - (If you generated an auth key earlier, **revoke it** — it's not used.)

2. **Confirm HTTPS is on** (you said it is). Admin console → DNS → HTTPS Certificates
   enabled. This gives her a no-warning cert and makes GPS/camera work.

3. **Tell her the treatments PIN** if she'll edit prices: `9245` (in `.env`).
   The setup sheet already unlocks the editor for her automatically.

---

## Part B — things on YOUR LAPTOP (one-time)

1. **Restart the Sheila server** to pick up the new `:8443` tunnel setting:
   - Stop any running `node server.js` for Sheila.
   - Start it again (or run `start.cmd`). On boot it runs
     `tailscale serve --https=8443 → 3001` and prints the `Public URL` line ending
     in `:8443`. Confirm it matches the address above.
   - Your **Receipts** app keeps `:443` — they no longer fight.

2. **Enable auto-start** so the server is always running when the laptop's on:

       powershell -ExecutionPolicy Bypass -File .\enable-autostart.ps1

   This drops a "Sheila Server.lnk" in your Startup folder. (Disable later by
   deleting that shortcut from `shell:startup`.)

3. **Stop the laptop sleeping** (or it's unreachable while asleep):
   Settings → System → Power → Screen and sleep →
   *"When plugged in, put my device to sleep"* = **Never**.

4. **Sanity check** the tunnel from your own phone/browser:
   open `https://desktop-rh302uu.tail999c81.ts.net:8443/api/health` —
   you want `{"ok":true,"excelReady":true,...}`.
   If `excelReady` is false, openpyxl/Python isn't ready (see README).

---

## Part C — what you send HER

Send her **`SHEILA-PHONE-SETUP.html`** (open it on her phone, or print it).
It has the three QR codes hardcoded to her address:
1. Install Tailscale, 2. Open the app + Add to Home Screen, 3. Treatments editor.

Then walk her through: install Tailscale → **Sign in with Apple** → accept your
invite email → toggle On → scan the app QR → Add to Home Screen → allow Camera +
Location. The sheet itself spells this out, so it's mostly a backstop.

---

## TEST THIS YOURSELF FIRST (you're 700km away — prove the path before she relies on it)

The invite route puts her phone on your tailnet as a *shared-in* device. Tailscale's
HTTPS cert sometimes does NOT cover shared devices — if so she'd get a "Not Secure"
warning and GPS/camera may be blocked. **Unknown until a shared device tries it.**

Before she sets up, replicate her exact path with a device YOU control:

1. Invite a **second account you own** (or a spare phone) to your tailnet the same
   way you'll invite her (Users → Invite external users).
2. On that device: install Tailscale → sign in with that second account → accept
   the invite → toggle On.
3. Open `https://desktop-rh302uu.tail999c81.ts.net:8443` on it.
   - **Loads clean, camera + GPS work** → her path is proven. Ship it.
   - **"Not Secure" warning** → the shared-device cert issue is real. Fix options:
     a) Switch her to your-account sign-in (no shared-device cert gap), OR
     b) Use Tailscale **node sharing** instead of user invite, OR
     c) Have her install the cert (see below) — last resort, needs you on a call.

## If she gets a "Not Secure" warning (cert fix, walk her through on a video call)

Her sheet tells her to call you, not to attempt this alone. On the call:

1. In Safari on her phone, open `https://desktop-rh302uu.tail999c81.ts.net:8443/cert`
   — it downloads `Sheila.crt`.
2. Settings → **Profile Downloaded** → Install (top right) → enter passcode → Install.
3. Settings → General → About → **Certificate Trust Settings** → toggle **ON** for
   the Sheila certificate.
4. Reopen the app — warning gone, camera/GPS work.

This is fiddly; it's exactly why her sheet says "call Cormac" rather than asking
her to do it solo.

## Honest limitations / things to watch

- **Her data lives on YOUR laptop** (`her-expenses.xlsx`). If you want her to have
  it, you'll need to share the file out separately (e.g. drop it in OneDrive).
- **If your laptop is off, nothing files** until it's back — her phone holds the
  queue safely, but the Excel file only updates when the server is reachable.
- **GPS** depends on the trusted Tailscale cert; it should work over `:8443` HTTPS.
  It did NOT capture during the local same-Wi-Fi test (self-signed cert) — worth
  confirming once she's on Tailscale.
- **The In (income) flow** hasn't been tested on a real phone yet — only Out.
- If you ever change machines or your `.ts.net` name changes, her QR/address
  changes too and you'll need to re-send the setup sheet.
