# Lumber Load — iOS app

A native iOS wrapper (via [Capacitor](https://capacitorjs.com/)) that combines both
Centerbeam web tools into one app with a bottom tab bar:

- 📋 **Tally** — `centerbeam_tally_recommender.html`
- 🧮 **Planner** — `centerbeam_layout_planner.html`

Each tool runs inside its own iframe, fully offline. The repo-root HTML files stay
the single source of truth; `scripts/sync-www.mjs` mirrors them into `www/` (as
`tally.html` / `planner.html`) before every build.

**Native haptics:** the two web apps call `@capacitor/haptics` (real Taptic Engine)
via a small `Haptic` helper that resolves the plugin through the iframe boundary and
no-ops on the plain web. `cap sync` on the Mac registers the plugin natively — no
extra steps. Toggles/taps give a selection tick, add/remove/clear a light impact, and
a perfect Solve / a found Recommendation give a success notification.

```
app/
  capacitor.config.json   # appId, appName, iOS settings
  package.json            # scripts + Capacitor deps
  scripts/sync-www.mjs    # copies the 2 root HTML apps into www/
  www/
    index.html            # the tab-bar SHELL (committed; not synced)
    tally.html            # synced copy (gitignored)
    planner.html          # synced copy (gitignored)
  ios/                    # generated on a Mac by `cap add ios` (Xcode project)
```

---

## Build & run on your iPhone

### What can be done on Windows (already set up here)
- Project config, the tab shell, dependency install, and `www/` sync.
- `npm install` and `npm run sync-www` work on Windows.

### What requires a Mac (Xcode is macOS-only)
Adding the iOS platform runs CocoaPods, and the compile/sign/run step needs Xcode.

On a Mac with **Xcode** + **CocoaPods** installed, from this `app/` folder:

```bash
npm install
npm run cap:add-ios     # syncs www/ then `cap add ios`  (run once)
npm run cap:open        # opens ios/App/App.xcworkspace in Xcode
```

In Xcode:
1. Select the **App** target → **Signing & Capabilities**.
2. **Team** → sign in with your **free Apple ID** (Add Account…). Personal team is fine.
3. Set a unique **Bundle Identifier** if `com.kenpaine.lumberload` is taken.
4. Plug in your iPhone, pick it as the run destination, press **▶ Run**.
5. On the phone: **Settings → General → VPN & Device Management** → trust your developer cert.

> Free Apple ID limits: the app is signed for **7 days** (re-run from Xcode to renew),
> max 3 sideloaded apps. No paid Developer account or TestFlight needed for personal use.

### After editing either web app
Re-mirror the root HTML into the app and push to the native project:

```bash
npm run cap:sync        # sync-www + `cap sync ios`
```
(no need to re-run `cap add ios`).

---

## No Mac handy?
Build the IPA once in the cloud (Codemagic free tier, or a GitHub Actions
`macos-latest` runner), then install/re-sign from Windows with **Sideloadly** or
**AltStore** — same free 7-day signing, done over Wi-Fi.

## TODO
- App icon + splash: drop a 1024×1024 PNG and run `@capacitor/assets` on the Mac.
