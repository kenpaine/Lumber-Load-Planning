# Packaging Lumber Loader as Native Store Apps

A build-and-submit checklist for shipping `lumber_loader.html` as
native apps in the **Apple App Store, Google Play, Microsoft Store, and Mac App
Store**.

**Architecture:** keep the existing single-file web app as the shared source of
truth and wrap it — **[Capacitor](https://capacitorjs.com)** for mobile (iOS +
Android), **[Tauri v2](https://tauri.app)** for desktop (Windows + macOS). No
rewrite; every store wraps the same `www/index.html`.

> The iOS Capacitor shell already exists in `app/` (see `app/README.md`). The
> Android, Windows, and Mac builds follow the same pattern — point `webDir` at
> the built `www/` folder which mirrors `lumber_loader.html`.

> The app is pure client-side computation with **no backend, no dependencies, and
> no data collection** — which keeps every store's privacy/data forms trivial
> ("no data collected") and review straightforward.

---

## Phase 0 — One-time foundation (shared by all stores)

**Accounts** (start early — Apple enrollment can take a day or two):
- [ ] Apple Developer Program — **$99/yr** (covers iOS **and** Mac App Store)
- [ ] Google Play Console — **$25 one-time**
- [ ] Microsoft Partner Center (individual) — **~$19 one-time**

**Identity & assets** (define once, reuse everywhere):
- [ ] App name: *Lumber Loader*
- [ ] Bundle/App ID (reverse-DNS, identical across stores): `com.kenpaine.lumberload` (already set in `app/capacitor.config.json`)
- [ ] Master icon: one **1024×1024 PNG** → generate all sizes (`capacitor-assets` / `tauri icon`)
- [ ] **Privacy policy URL** — required by Apple & Google. A one-paragraph "no data
      collected, runs fully offline" page on GitHub Pages is enough.
- [ ] Version scheme: semver (`1.0.0`) + an integer build number bumped every upload

**Tooling:**
- [ ] Node ✅ · Android Studio + SDK (Windows) · Xcode (Mac) · Rust for the desktop shell (`winget install Rustlang.Rustup`)
- [ ] Capacitor: `npm i @capacitor/core @capacitor/cli`; copy `lumber_loader.html` → `www/index.html` as the shared `webDir` (the `app/` directory already does this via `sync-www` in `app/package.json`)

---

## Store 1 — Google Play (Android) · build on Windows

**Build**
- [ ] `npx cap add android` → `npx cap sync`
- [ ] In Android Studio: set `applicationId`, `versionCode`/`versionName`, adaptive icon
- [ ] Create an **upload keystore** (`keytool -genkey …`) and enable **Play App Signing**
- [ ] Build a signed release **AAB** (`./gradlew bundleRelease`)

**Submit**
- [ ] Play Console → create app → store listing (title, short/full description,
      **phone screenshots**, feature graphic **1024×500**, icon **512×512**)
- [ ] Content rating questionnaire (→ Everyone) · **Data safety** form (→ "No data
      collected") · Target audience · Privacy policy URL
- [ ] Upload AAB to **Internal testing** → verify → promote to **Production** (free, choose countries)
- [ ] Submit for review *(typically hours–2 days)*

---

## Store 2 — Apple App Store (iOS) · build on the Mac

**Build**
- [ ] On Mac: install Xcode + CLT, pull the repo, `npx cap add ios` → `npx cap sync ios`
- [ ] Apple Developer → register the **App ID** (bundle id) → create the app in
      **App Store Connect** (name, SKU, bundle id)
- [ ] Xcode: set bundle id, version/build, **automatic signing** (your team), app icon, launch screen
- [ ] Product → **Archive** → Validate → **Upload** (Xcode Organizer or Transporter)

**Submit**
- [ ] App Store Connect listing: description, keywords, support URL, **privacy policy
      URL**, **screenshots** (6.7" iPhone required; add iPad sizes only if you support iPad)
- [ ] **App Privacy** label → "Data Not Collected" · Age rating → 4+ · Price → Free
- [ ] Submit for review *(~1–3 days)*

> ⚠️ **The one real risk to plan for:** Apple **Guideline 4.2 (minimum
> functionality)** — they sometimes reject thin web-wrapper apps that feel like
> "just a website." Mitigate by making it feel native: works **fully offline**, no
> browser chrome, app icon/splash, and ideally a small platform touch (e.g. haptic
> on a result, share-sheet export of a tally). A genuinely useful offline
> calculator usually passes, but build it app-like, not webview-like.

---

## Store 3 — Microsoft Store (Windows) · build on Windows (Tauri)

**Build**
- [ ] Install Rust; `npm create tauri-app` (or add Tauri), point `frontendDist` at the web folder
- [ ] `tauri.conf.json`: productName, identifier, version, window size; `npm run tauri icon` from the 1024 PNG
- [ ] `npm run tauri build` → produces `.msi`/`.exe`; configure the **MSIX** target (required for the Store)

**Submit**
- [ ] Partner Center → reserve app name → new submission
- [ ] Upload the **MSIX** package · store listing (description, **screenshots
      ≥1366×768**, icon) · **IARC age rating** · Markets · Price → Free
- [ ] Submit for certification *(hours–1 day)*

> Store-distributed MSIX is signed by Microsoft. For *direct* `.exe` downloads you'd
> want your own code-signing cert (~$70–200/yr) to avoid the SmartScreen warning — optional.

---

## Store 4 — Mac App Store (macOS) · build on the Mac (Tauri)

**Build**
- [ ] On Mac: `npm run tauri build` → `.app`/`.dmg`
- [ ] For the **store**: add **App Sandbox entitlements**, sign with **3rd-Party Mac
      Developer** certs + provisioning profile, package as **`.pkg`**

**Submit**
- [ ] App Store Connect → create the **macOS** app entry · listing + **Mac
      screenshots** · privacy label · price
- [ ] Upload via **Transporter** → submit for review

> 🪧 **Strong recommendation:** for Mac, **skip the Mac App Store** and instead
> **notarize a direct `.dmg`** (Developer ID cert + `notarytool`). It's far less
> fiddly than the sandbox/entitlements dance, and Mac users routinely install
> notarized `.dmg`s directly. Only do the Mac App Store if store *discoverability*
> specifically matters to you.

---

## Reality check on effort

- **Easiest first win:** Google Play (build right here on Windows, lenient review). Do it first to learn the rhythm.
- **Biggest grind across all stores:** screenshots per device class, and the listing copy — not the code.
- **Your code barely changes** — every store wraps the same `www/index.html`.

---

## Cross-cutting prep (do once, reuse for every store)

- [ ] **Privacy policy** page (hosted, e.g. GitHub Pages) — "collects no data, runs fully offline"
- [ ] **App icon** master 1024×1024 → all platform sizes
- [ ] **Screenshots** per platform/device class (the main grind)
- [ ] **Content/age ratings** → Everyone / 4+
- [ ] **Data-safety / privacy labels** → "No data collected" everywhere (true for this app)
- [ ] Consistent **bundle ID** + **version/build** discipline across stores
