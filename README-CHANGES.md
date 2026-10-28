# Recent changes — deploy & verification guide

> **Files changed this round:** `LoanManagerV2.gs`, `Validaciones.gs`
> `Reactivar.gs` and `SignatureData.gs` are **unchanged** — no need to paste them.

---

## What's new

### 1 · Balance fix in WhatsApp messages and the Payments sheet (90-day loans)

90-day loans have 3 installments. The "Total a Pagar" column was only storing the principal, not the full installment plan total, so WhatsApp confirmations and the "Saldo Posterior" column in the *Pagos* sheet were showing the wrong remaining balance.

**Fixed in `LoanManagerV2.gs`:**

| Location | What changed |
|---|---|
| `loanWaStatus_` | Now sums the real installment balances from the *Cuotas* sheet instead of reading "Total a Pagar" |
| `paymentBalanceFormula_` | The *Pagos* sheet formula now prefers the "Total del PLAN" column (sum of all installments) when it exists |

### 2 · Payment emails and receipts now include installment breakdown

The payment confirmation email and the attached PDF now show the same detail as the WhatsApp message: installment-by-installment status, partial payment amounts, due dates, late fees, and frozen/payment-plan states.

**Updated in `LoanManagerV2.gs`:**

| Function | What was added |
|---|---|
| `cuotasHtmlTable_` | New — renders an HTML table with the status of each installment (for email bodies) |
| `receiptHtml_` | Shows installment breakdown, late fees, and real balance (optional `s` parameter) |
| `makePaymentReceiptFile_` | PDF now has one row per installment, outstanding balance, and late fee |
| `recordPayment_` | Computes `loanWaStatus_` once and shares it with WhatsApp, email, and PDF |
| `sendPaymentReceiptForRow_` | Re-sent receipts now use `loanWaStatus_` for the correct balance |

### 3 · Application form: edit email in "Your details" (Step 2)

Returning clients saw their name, phone, and address pre-filled in Step 2 but had no way to see or update their email without going back to Step 1.

**Added in `Validaciones.gs`:**

- Step 2 ("Tus datos") now displays the registered email below the address field.
- A **"Cambiar"** button makes the field editable, pre-filled with the current address.
- When the client changes the email and advances to Step 3, the format is validated and the value is synced back to the form before submission.
- No backend changes needed — `submitIntakeSmart` already called `updateClienteEmail_` whenever the submitted email differed from the stored one.

### 4 · Application form: references are always editable

Previously, if a returning client had saved references, Step 3 was skipped automatically with no way to update a changed phone number or contact.

**Changed in `Validaciones.gs`:**

- The references form is always shown — never auto-skipped.
- If the client has references on file, all 6 fields are pre-filled with the stored values and a note reads **"Referencias en archivo — podés actualizarlas si cambiaron."**
- On submit, the server always saves the form values (it already did — only the frontend behaviour changed).

### 5 · Welcome card with client profile summary

The plain "¡Hola de nuevo, Denis! Ya te tenemos registrado (C-0028)…" text box was replaced by a compact profile card that fits in the same space and shows actionable information at a glance:

```
¡Hola de nuevo, Denis Bell!   Nivel 2 ★★☆☆☆
💰 $150,000   ✓ 4 payments   No active loans
⭐ Pay on time and your next limit goes up to $200,000.
```

**Row 1 — name + star rating:**
- 1–5 stars based on the client's position in the graduation ladder (`r.levelNum / r.maxLevelNum`)
- Colour: gold (normal) · green (maximum level) · red (outstanding debt)
- The label "Nivel X" or "Inicial" is shown next to the stars

**Row 2 — 3 status chips:**

| Chip | Data source |
|---|---|
| 💰 Available limit | `r.maxAmountFmt` (computed by `graduationLevelFor_`) |
| Payment history | `r.settledCount`, `r.hasArrears`, `r.levelDropped` |
| Active loans | `r.activeLoans` (new — calls `clientLoanCount_` server-side) |

**Row 3 — early-repayment incentive line** (uses `r.nextMax`):

| Situation | Message |
|---|---|
| Outstanding debt | ⚠ Settle your debt to recover your limit. |
| Maximum level | 🏆 You have the highest limit. Keep paying on time to keep it. |
| Level dropped / reset by late payment | ↑ One on-time payment recovers a level → $X. |
| Good standing, not yet at max | ⭐ Pay on time and your next limit goes up to $X. |

The client ID (C-0028) is no longer shown to the borrower.

**Backend change (`Validaciones.gs` — `lookupClienteForIntake`):** `res.activeLoans = clientLoanCount_(cByDni.id)` is added to the return object for found clients. `clientLoanCount_` already existed in the same file.

---

## How to apply the changes

### 0) Back up first

**File ▸ Make a copy** before touching anything, or confirm **File ▸ Version history** is available to roll back if needed.

### 1) Paste the updated code

1. Open **Extensions ▸ Apps Script**.
2. Replace the full contents of **`LoanManagerV2.gs`** and **`Validaciones.gs`** with the updated files.
3. Save (💾).
4. Return to the sheet and **reload the browser tab** so the "Gestor de Préstamos" menu rebuilds.

> No need to run `reactivarFuncionalidad` or `instalarTokensDeAcceso` — this round adds no new columns or sheets.

### 2) Publish a new web app version

**Deploy ▸ Manage deployments ▸ ✏️ (edit active) ▸ Version: New version ▸ Deploy.**

The form changes (welcome card, editable email, pre-filled references) only reach users after a new version is published. The `…/exec` URL stays the same.

---

## Verification

### A) WhatsApp balance / Saldo Posterior (fixes 1 & 2)

- [ ] Open loan L-0015 (or any 90-day multi-installment loan) in *Pagos Atrasados*.
- [ ] Trigger the WhatsApp send → the message should show the correct balance (sum of outstanding installments, not just the principal).
- [ ] In the *Pagos* sheet, the **"Saldo Posterior"** column for P-0044 should show $800,000 (not $50,000).

### B) Email and PDF with installment breakdown (fix 3)

- [ ] Record a partial payment on a 90-day loan → the confirmation email should include a table with the status of each installment.
- [ ] Tick **"Enviar recibo"** for that payment → the email and PDF attachment show the breakdown and the correct balance.
- [ ] On a 15- or 30-day loan (single installment), the email shows only "Saldo pendiente" — no installment table.

### C) Editable email in Step 2

- [ ] Open the application form (no query parameters) and enter the email + DNI of an existing client.
- [ ] Advance to Step 2 → below "Dirección" a **"Correo electrónico"** row appears showing the registered email.
- [ ] Click **"Cambiar"** → the field becomes editable, pre-filled with the current email.
- [ ] Change the email and advance to Step 3 → an invalid format shows an inline error and blocks progress.
- [ ] Submit with the changed email → in the *Clientes* sheet, the client's email has been updated.

### D) Pre-filled editable references

- [ ] Client with saved references: on reaching Step 3, all 6 fields are pre-filled and the note "podés actualizarlas si cambiaron" is visible.
- [ ] Change a reference phone number and submit → in *Clientes*, the phone number has been updated.
- [ ] Client with no saved references: Step 3 is blank as before.

### E) Welcome card

- [ ] Mid-level client: card shows name + `Nivel X ★★★☆☆` + 3 chips + the "⭐ Pay on time…" incentive line.
- [ ] Client with outstanding debt: stars in red, "⚠ Deuda vencida" chip, red incentive line.
- [ ] Client at maximum level: stars in green, "🏆 Tenés el límite máximo…" message.
- [ ] The client ID (C-XXXX) no longer appears anywhere in Step 2.

### F) HTML preview (optional)

Open `simulation/welcome-card-preview.html` in a browser to see all 7 card variants without connecting to Google Sheets.

---

## Quick run order

```
Backup → paste LoanManagerV2.gs + Validaciones.gs → Save → reload tab
Deploy ▸ New version  (so the updated form reaches users)
```
