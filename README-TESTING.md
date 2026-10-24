# Testing guide — this round's new features

How to test, **on a COPY** of the system (never in production), everything added in this round:

1. **Third‑party account** (money goes to someone else's account) — form, validations, holder consent, email + contract to the holder, row coloring, lender notification.
2. **Device capture** on loan application → **Dispositivos** sheet.
3. **C‑2 security** — per‑loan token in the links + anti‑brute‑force lockout on the self‑service forms (sign / balance / freeze / plan).
4. **Cloudflare Turnstile** (anti‑bot) on every public form.

> In‑app labels (menus, settings, sheet/column names, messages) are quoted **in Spanish** on purpose — that's exactly what you'll see, because the app's UI is Spanish.
>
> Files changed this round: **`LoanManagerV2.gs`**, **`Validaciones.gs`**, **`Reactivar.gs`**.
> Unchanged: `VerificacionDNI.gs`, `SignatureData.gs`, `Migracion.gs`, `RepararPagos.gs`.

---

## 0) Create the copy and paste the code

1. In Google Sheets: **File ▸ Make a copy** of the production sheet (the data and the bound Apps Script project are copied, isolated from production). *Do not test on the real sheet.*
2. In the copy: **Extensions ▸ Apps Script**. Paste the updated contents of **`LoanManagerV2.gs`**, **`Validaciones.gs`** and **`Reactivar.gs`** (replace each full file) and **Save** (💾).
3. Go back to the sheet and **reload the browser tab** so the **"Gestor de Préstamos"** menu is rebuilt.

---

## 1) Install (once, IN THIS ORDER)

| # | Where | Action | Installs |
|---|-------|--------|----------|
| 1 | Menu **Gestor de Préstamos** | **"🆕 Aplicar novedades (columnas y ajustes)"** | Adds the new settings (token / Turnstile) to **Configuración**, non‑destructively |
| 2 | Apps Script editor ▸ Run | **`instalarTokensDeAcceso`** | Creates the **`Token de acceso` / `Token emitido`** columns in *Prestatarios* and mints one token per existing loan |
| 3 | Apps Script editor ▸ Run | **`reactivarFuncionalidad`** | Third‑party columns + green/red row coloring in *Nuevos Prestatarios*, and the **"Enlace de Firma"** formula now with `&t=` |
| 4 | Apps Script editor ▸ Run | **`activarDisparadores`** | Reinstalls the `onEdit` trigger and the daily tasks (copies **lose** installable triggers) |

> Order 2→3 matters: `instalarTokensDeAcceso` creates the token column **before** `reactivarFuncionalidad` writes the link formula that references it.
> Authorize the permissions when the editor asks (Sheets / Drive / Gmail / UrlFetch).
> **Do not run `setupConfirm`** ("① Configurar / reconstruir hojas") on a copy that has data: it rebuilds the sheets and is destructive. Use `reactivarFuncionalidad` (non‑destructive).

---

## 2) Configuration (the "Configuración" sheet)

| Setting | Purpose | Value for testing |
|---------|---------|-------------------|
| **"URL de la app web (enlaces a clientes)"** | Make the links point to **THIS copy**, not production | Paste the copy's `…/exec` (see step 3) |
| **"Aceptar solicitudes de préstamo (SÍ/NO)"** | Enable the application form | `SÍ` |
| **"Vigencia del enlace (minutos)"** | Link‑token lifetime (C‑2) | `1440` (set `1` to test expiry) |
| **"Exigir token en enlaces (SÍ/NO)"** | C‑2 on/off switch (leave on) | `SÍ` (set `NO` only to disable) |
| **"Intentos antes de bloquear (enlaces)"** / **"Bloqueo por intentos (minutos)"** | Anti‑brute‑force | `8` / `15` |
| **"Capturar ubicación en la firma (SÍ/NO)"** | GPS/IP in device capture (also on the application) | `SÍ` or `NO` |
| **"Turnstile Site Key"** / **"Turnstile Secret Key"** | Turn the anti‑bot on (blank = off) | Paste the keys from the Cloudflare Turnstile dashboard |

---

## 3) Publish the copy's web app

1. Editor ▸ **Deploy ▸ New deployment** ▸ type **Web app** ▸ *Execute as: Me*, *Who has access: Anyone* ▸ **Deploy**. Copy the **`…/exec`** URL.
2. Paste that URL into **Configuración ▸ "URL de la app web (enlaces a clientes)"**.
3. For fast iteration use the **`…/dev`** URL (always runs the latest code); the `…/exec` URL needs a *New version* whenever you change code.

---

## 4) Tests per feature

### A) Third‑party account

- [ ] **A1 · Own account:** Step 4 → *Transferencia* → enter an alias → "¿La cuenta está a tu nombre?" = **Sí** → submits. In *Nuevos Prestatarios*: `¿Cuenta propia? = Sí`, **GREEN row**. The lender email shows *"Cuenta de destino: <alias> — a nombre del solicitante"*.
- [ ] **A2 · Cash (Efectivo):** choose *Efectivo* → the whole account/holder block hides and it submits without asking for an account.
- [ ] **A3 · Third‑party incomplete (blocked):** *Transferencia* + **No** with no data → **"Siguiente"** won't advance (red errors). If you bypass the JS, the server rejects with the **V‑36** message.
- [ ] **A4 · Holder must be a different person:** set the holder's DNI / email / phone **equal** to the applicant's → rejected (*"…no puede ser el mismo que el del solicitante"*). The **name** may match.
- [ ] **A5 · Third‑party complete:** holder full name (≥10), DNI, email, phone (WhatsApp) + **DNI photos front and back** → submits. In *Nuevos Prestatarios*: **RED row**; columns `¿Cuenta propia? = No (tercero)`, `Titular de la cuenta`, `DNI/Correo/Teléfono del titular`, `DNI titular (frente/dorso)` (Drive links), a **`Titular acepta?`** checkbox (unchecked) and **`Avisar al titular 📲`** (WhatsApp link). The lender email carries the holder block + the WhatsApp link + **"⚠ Acción requerida"**.
- [ ] **A6 · "Titular acepta?" gate (cannot be bypassed):** tick **`Verificado?`** WITHOUT ticking **`Titular acepta?`** → it bounces with the V‑36 message and unticks. Tick **`Titular acepta?`** and then **`Verificado?`** → the loan moves to *Prestatarios*. (Confirm **"Anular límites"** does NOT bypass this gate.)
- [ ] **A7 · Email + contract to the holder:** on approval, the **applicant** receives the contract; the **holder** receives *"Sos el titular de la cuenta…"* with the **PDF attached** and **no** "Ver el contrato" button. In the PDF, section **3 · "Medio de pago y cuenta de destino"** highlights payment method + alias + holder name/DNI/email/phone, and the "the account used is yours" clause is replaced by the third‑party authorization clause.

### B) Device capture → "Dispositivos" sheet

- [ ] **B1:** submit any application (own or third‑party). Open the **Dispositivos** sheet → a new row with `ID Préstamo = (solicitud)`, client ID, DNI, name, email, **SHA‑256 fingerprint**, device/OS, browser, screen, cores, memory, languages, timezone, IP, etc.
- [ ] **B2:** with *"Capturar ubicación en la firma" = SÍ* the browser asks for location on submit and the row has GPS/IP‑location; with **NO**, there's no prompt and only IP + fingerprint are stored.

### C) C‑2 security (token + lockout)

- [ ] **C1 · Legit flow:** approve an application; the "Revisar y firmar" email carries a `?page=firmar&loan=…&t=…` link. Open it, enter the **correct DNI** → the contract loads and can be signed.
- [ ] **C2:** in *Prestatarios*, the **"Enlace de Firma"** column shows a URL that includes `&t=`.
- [ ] **C3 · Enumeration blocked:** open `?page=firmar&loan=L-00XX` **without** `&t=` (or with a wrong `t=`) + correct DNI → rejected (*"Datos no encontrados o el enlace venció…"*); no PII/contract shown and nothing can be signed.
- [ ] **C4 · Expiry:** set *"Vigencia del enlace (minutos)" = 1*, wait >1 min, open a valid link → rejected as expired. Menu **"✉ Reenviar enlace de firma (fila seleccionada)"** → issues a fresh working link and **kills the old one**. Restore the setting to `1440`.
- [ ] **C5 · Lockout (throttle):** on the signing or balance page, make **8** attempts with a wrong DNI/token for the same loan → the **9th** shows *"Demasiados intentos…"* for 15 min.
- [ ] **C6 · Balance:** `?page=saldo` → loan ID + correct DNI → balance shows (no token needed); 8 failed attempts → lockout.
- [ ] **C7 · Freeze / plan:** the *congelar*/*plan* WhatsApp links include `&t=`; opening them + DNI works; without `t=` they're rejected.

### D) Cloudflare Turnstile (anti‑bot)

- [ ] **D1 · No keys:** with *Turnstile Site/Secret Key* blank, **all forms work as before** (no widget shown).
- [ ] **D2 · With keys:** paste Site + Secret Key, **redeploy / use `…/dev`**. Every public form (application, balance, freeze, plan, sign) shows the **Turnstile widget**; completing it lets the submit proceed and the server validates the token. On a failed verification the widget resets so the user can retry.
- [ ] **D3 · Caveat (iframe):** the app runs inside a `*.googleusercontent.com` iframe. In the **Cloudflare Turnstile** dashboard add that domain (and/or `script.google.com`) to the widget's **allowed hostnames**. Test **one** form end‑to‑end right after adding the keys; if the widget doesn't render or `siteverify` reports a hostname mismatch, adjust the widget's domains.

---

## 5) Kill switches / rollback

- **C‑2 (token):** Configuración ▸ *"Exigir token en enlaces (SÍ/NO)" = NO* → links go back to DNI‑only validation (no code change). The anti‑brute‑force lockout stays active.
- **Turnstile:** clear *"Turnstile Site Key"* and *"Turnstile Secret Key"* → disabled on all forms.
- **Location/GPS on the application:** *"Capturar ubicación en la firma (SÍ/NO)" = NO*.
- **Errors:** any failure is logged to the **Errores** sheet (e.g. `verifyTurnstile_`, `submitIntakeSmart:dispositivo`). Check it if something doesn't behave as expected.

## Quick run checklist

```
Copy the sheet → paste LoanManagerV2.gs / Validaciones.gs / Reactivar.gs → Save → reload
Menu:    🆕 Aplicar novedades (columnas y ajustes)
Editor:  instalarTokensDeAcceso   →   reactivarFuncionalidad   →   activarDisparadores
Config:  "URL de la app web" (the copy's URL) [+ Turnstile keys if testing]
Deploy:  Deploy ▸ New deployment ▸ Web app  → paste the …/exec into the "URL de la app web" setting
```
