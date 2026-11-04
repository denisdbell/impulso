# Deployment guide — Impulso Crédito

Changes from this work session, and exactly how to roll them out.
**Scope:** 3 Apps Script files (Google Sheets) + 1 Python desktop file.

> ⚠️ **Before you paste:** in the Apps Script editor, copy each current file's contents to a scratch
> doc first (simple rollback). Changes affect **new** loans/applications going forward; existing
> loans are untouched except the explicit repair steps below.

---

## What changed (summary)

| File | What |
|---|---|
| **LoanManagerV2.gs** | ① Welcome email to first-time borrowers (daily 18:00 job + admin report, once per person). ② Frozen-loan **mora over-charge fix** (`computeOutstanding_`). ③ Loans no longer become **implicit plans** by amount — schedule follows the loan's **term** (signing, contract PDF, test data). ④ Panel `withdrawableStats_` date-column lookup hardened. |
| **Reactivar.gs** | ① Signing-link column now includes the **access token** `&t=…` (`signingLinkFormula_` callers). ② Approval builds the cuota schedule from the **term**, not the amount. |
| **Validaciones.gs** | Loan terms by amount: **exactly $300,000 → 30 days; above $300,000 → 90 days** (server rule + intake-form mirror + tier help text). |
| **bcra/report_gui.py** | New **"Loans by month"** dashboard page (month/year/status filters, totals, "Exclude overdue"). Desktop app only — not Apps Script. |

---

## Part A — Apps Script (Google Sheets)

1. Open the spreadsheet → **Extensions ▸ Apps Script**.
2. For each file below: open it in the editor, **select all → paste the updated contents → Save**.
   - `LoanManagerV2.gs`
   - `Reactivar.gs`
   - `Validaciones.gs`
3. **Timezone** (needed so the welcome job fires at 6 PM Argentina): click **⚙ Project Settings**,
   enable **"Show appsscript.json"**, open `appsscript.json`, confirm:
   ```json
   { "timeZone": "America/Argentina/Buenos_Aires", "runtimeVersion": "V8" }
   ```
4. Reload the spreadsheet tab (so the "Gestor de Préstamos" menu reloads).

---

## Part B — Run these in the spreadsheet, in order

Everything below is idempotent (safe to re-run) and does **not** modify loan data except where noted.

1. **Reinstall live formulas + signing-link tokens** — editor ▸ Run **`reactivarFuncionalidad`**.
   Rewrites the **"Enlace de Firma"** column on every row to include `&t=<token>`, and rebuilds the
   Panel. *(Optional: Run **`instalarTokensDeAcceso`** first to fill the 1 loan missing a token.)*

2. **Recompute balances / mora / Panel** — menu **Gestor de Préstamos ▸ ⑤ Actualizar saldos y
   resumen**. This applies the **frozen-mora fix** (e.g. L-0034 "Joana" drops $67,500 → **$22,500**,
   Saldo con Mora $167,500 → **$122,500**) and un-breaks the **Panel** (fills "Efectivo disponible",
   fixes the "Ganancia real" label).

3. **Install the welcome email job** — menu **Gestor de Préstamos ▸ 👋 Bienvenida a nuevos
   prestatarios (automática) ▸ ⏰ Instalar / actualizar envío diario**. Adds the **"Bienvenida
   enviada"** column to *Clientes*, marks the ~39 existing clients as `preexistente` (no mass blast),
   and installs the **18:00 daily trigger**. First real run welcomes recent first-timers (first
   signed loan ≤ 7 days) and emails a recap to *Correo para resumen de avisos*.
   - To send immediately instead of waiting for 6 PM: **✉ Enviar bienvenidas ahora**.

4. **Repair loan L-0089** (it was wrongly split into a 4-installment $600k "plan"). Paste this into
   the editor, **Run `arreglarL0089`**, then re-run **⑤ Actualizar**:
   ```js
   function arreglarL0089() {
     var L = findLoanById_('L-0089');
     // reemplaza las 4 cuotas por 1 sola de $450.000 que vence a los 30 días
     cuotasForLoan_(getSS_().getSheetByName('Cuotas'), 'L-0089', L.loanDate, 450000, 1, [30]);
   }
   ```
   Result: 1 payment of **$450,000 due 2026-12-03**, status returns to **ACTIVO** (no longer EN PLAN).

---

## Part C — Desktop dashboard (`bcra/report_gui.py`)

No "deploy" step — it's a local Python file. To use the new **Loans by month** page:
```bash
cd bcra && python3 report_gui.py --xlsx out/_impulso.xlsx
```
If you don't have service-account creds, open **Options ▸ uncheck "Download latest workbook on
reload"** so it loads the local workbook. Click **📅 Loans by month**.

---

## Verification checklist

- [ ] A loan's **Enlace de Firma** URL ends with `&t=…` (pick any row in *Prestatarios*).
- [ ] **Panel** shows "Efectivo disponible para prestar" ≈ **$109,991** (not blank) and a clean
      "Ganancia real (neto…)" label.
- [ ] **L-0034 (Joana)** in *Pagos Atrasados*: mora = **$22,500** (was $67,500).
- [ ] **L-0089** is **ACTIVO**, 1 cuota of **$450,000**.
- [ ] New application for **$300,000** in the web form offers **30 days**; **$300,001** offers **90 days**.
- [ ] *Clientes* has a **"Bienvenida enviada"** column; Triggers panel shows a daily
      `enviarBienvenidaPrestatarios` at 18:00.

---

## Not part of this deploy
`AUDIT-DATA-ACCURACY.md` and `bcra/analyze_impulso.py` were already modified before this session.
`welcome_email_preview.html` / `welcome_report_preview.html` are local email previews (not needed in
production).
