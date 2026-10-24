# Deploy guide — apply this round's changes to the existing Google Sheet

How to apply everything from this round of work to your **live** Google Sheet (the one exported as `impulso-v2 (25).xlsx`). This is an **in‑place** update — do the backup step first.

> In‑app labels (menus, settings, columns, functions) are quoted **in Spanish** on purpose — that's what you'll see, because the UI is Spanish.
> Files changed this round: **`LoanManagerV2.gs`**, **`Validaciones.gs`**, **`Reactivar.gs`**.
> Unchanged: `VerificacionDNI.gs`, `SignatureData.gs`, `Migracion.gs`, `RepararPagos.gs`.

## What's included
1. **Cuenta de terceros** — full flow (columns already exist on your sheet) + the new **explicit "datos de la OTRA persona" copy** and "must be exact or the loan isn't approved" wording; DNI/email/phone of the holder must differ from the applicant; non‑bypassable `Titular acepta?` gate; holder email + contract; red/green row coloring.
2. **Captura de dispositivo** on each application → **Dispositivos** sheet (sheet already exists).
3. **C‑2 security (NEW on this sheet):** a per‑loan access **token** on every emailed/WhatsApp link + anti‑brute‑force **lockout** on the self‑service forms. Requires new columns + settings (steps 4–5 below).
4. **Anti‑bot:** Cloudflare Turnstile was **removed** (it can't run in the Apps Script iframe) and replaced with a free in‑iframe **honeypot + dwell‑time** check on the application form. No action needed — your sheet has no Turnstile keys.
5. **Performance:** client‑side **photo compression** before upload, **batched** sheet writes, cached Drive folder id — a slow submit (tens of seconds) drops to a few seconds.
6. **Fixes:** the 96%‑stuck submit (file‑form must be the only parameter), device‑capture GPS timeout, lender‑notification fields, contract shows payment method + holder details.

---

## 0) Back up first
- **File ▸ Make a copy** (keep it aside), or confirm **File ▸ Version history** is available to roll back. This update is non‑destructive, but always keep a restore point for production.

## 1) Paste the updated code
1. **Extensions ▸ Apps Script**. Replace the full contents of **`LoanManagerV2.gs`**, **`Validaciones.gs`**, **`Reactivar.gs`** with the updated files. **Save** (💾).
2. Back on the sheet, **reload the browser tab** so the **"Gestor de Préstamos"** menu rebuilds.

## 2) Add the new settings (non‑destructive)
Menu **Gestor de Préstamos ▸ "🆕 Aplicar novedades (columnas y ajustes)"**. This adds the new rows to **Configuración** (token/lockout) without touching your data: `Vigencia del enlace (minutos)`, `Exigir token en enlaces (SÍ/NO)`, `Intentos antes de bloquear (enlaces)`, `Bloqueo por intentos (minutos)`.

## 3) Install the C‑2 token columns + mint tokens
Apps Script editor ▸ pick **`instalarTokensDeAcceso`** from the function dropdown ▸ **Run** (authorize if asked). This creates **`Token de acceso`** / **`Token emitido`** in *Prestatarios* and mints a token for every existing loan. (Your sheet doesn't have these yet — required for the signing links.)

## 4) Re‑install live formulas / formatting
Apps Script editor ▸ Run **`reactivarFuncionalidad`**. Refreshes formulas/validations and makes the **"Enlace de Firma"** column tokenized (`&t=…`), plus the red/green row coloring in *Nuevos Prestatarios*.

> **Order matters:** step 3 (`instalarTokensDeAcceso`) must run **before** step 4, so the link formula can reference the token column.

## 5) Ensure the triggers are active
Apps Script editor ▸ Run **`activarDisparadores`** (idempotent — safe to re‑run). Installs/updates the `onEdit` trigger and daily tasks. (A live sheet usually has these already; running it again does no harm.)

> **Do NOT run `setupConfirm`** ("① Configurar / reconstruir hojas") — it rebuilds sheets and is destructive. `reactivarFuncionalidad` is the safe one.

## 6) Set the Configuración values
Open the **Configuración** sheet and set:

| Setting | Value | Why |
|--------|-------|-----|
| **"Vigencia del enlace (minutos)"** | `1440` (24 h) | Token/link lifetime. Resend (`✉ Reenviar enlace de firma`) rotates it and restarts the window. |
| **"Capturar ubicación en la firma (SÍ/NO)"** | `NO` (recommended) | GPS **cannot** work inside the Apps Script iframe; `NO` skips it (no prompt, faster submit) while still capturing device fingerprint + IP. |
| **"URL de la app web (enlaces a clientes)"** | the **direct `…/exec`** URL (see step 7) | ⚠️ Your sheet currently points to a **tinyurl** (`tinyurl.com/Impulso-credito`). Tokenized links add `?page=…&loan=…&t=…`; if the short link doesn't forward query parameters, those links break. Use the direct `…/exec` URL here (keep the tinyurl only as a public‑form shortcut). |
| **"Exigir token en enlaces (SÍ/NO)"** | `SÍ` | Keep the token requirement on. (`NO` = DNI‑only fallback.) |
| **"Aceptar solicitudes de préstamo (SÍ/NO)"** | `SI` | Keep accepting applications. |

(The lockout settings `Intentos antes de bloquear (enlaces)` and `Bloqueo por intentos (minutos)` can stay as added — defaults 8 / 15 apply if blank. Make sure "Intentos…" holds a **number** like `8`, not text.)

## 7) Redeploy the web app
**Deploy ▸ Manage deployments ▸ ✏️ (edit active) ▸ Version: New version ▸ Deploy.** Copy the **`…/exec`** URL and paste it into **"URL de la app web (enlaces a clientes)"** (step 6). The form changes (explicit third‑party copy, honeypot, photo compression, no Turnstile) only reach users after this new version. *(The `…/dev` URL always runs the latest saved code if you want to preview first.)*

---

## 8) Verify (quick smoke test)
- **Application speed:** submit a test application with photos — it should finish in a few seconds (not tens of seconds); you'll see "Optimizando las fotos…".
- **Third‑party clarity:** Step 4 → Transferencia → "No, es de otra persona" shows the bold **"Datos de la OTRA persona"** block with the red "must be exact or the loan isn't approved" note; entering the holder's DNI/email/phone equal to the applicant's is rejected.
- **Row color:** own‑account application row is **green**, third‑party is **red** in *Nuevos Prestatarios*; `Titular acepta?` must be ticked before approval.
- **Token link:** approve an application → the "Revisar y firmar" email link is `?page=firmar&loan=…&t=…`; open it + correct DNI → contract loads. Opening it **without** `&t=` is rejected.
- **No Turnstile:** the forms show no Cloudflare widget; the application still submits.
- **Device log:** a new row appears in **Dispositivos** (`ID Préstamo = (solicitud)`).

## Rollback / switches
- **C‑2 token:** `Exigir token en enlaces = NO` reverts links to DNI‑only (lockout stays on).
- **Location prompt:** `Capturar ubicación en la firma = NO` (already recommended above).
- **Errors:** anything that fails is logged to the **Errores** sheet.
- **Full rollback:** restore from the backup copy (step 0) or File ▸ Version history.

## Quick run order
```
Backup (Make a copy) → paste LoanManagerV2 / Validaciones / Reactivar → Save → reload
Menu:   🆕 Aplicar novedades (columnas y ajustes)
Editor: instalarTokensDeAcceso  →  reactivarFuncionalidad  →  activarDisparadores
Config: Vigencia del enlace=1440 · Capturar ubicación=NO · URL de la app web=…/exec · Exigir token=SÍ
Deploy: Manage deployments ▸ New version ▸ paste the …/exec into "URL de la app web"
```
