# Impulso Crédito — Data-Accuracy Audit & Repair Plan

**Date:** 2026-10-20
**Auditor:** Claude (Opus 4.8)
**Source:** production export `impulso-v2 (20).xlsx`
**Method:** every row of Prestatarios / Pagos / Cuotas / Resumen / Clientes extracted with `openpyxl` (cached values) and the financials independently recomputed in Python. Flagged rows were then drilled with the raw cells. Companion code audit: `AUDIT.md`.

> Scope note: the `.xlsx` is a snapshot. All **data repairs** run against the **live Google Sheet** via Apps Script (the export can't be edited back in). **Code fixes** go in the `.gs` files.

---

## Summary

The ledger-level math is **sound** — pricing, payment reconciliation, per-client aggregates, and cash all tie out exactly. The real exposure is the **installment (Cuotas) schedule**, which bills ~2× the loan total for 6 loans.

### ✅ Reconciled exactly (verified)
- **Pricing — 61/61 loans:** `Interés = Capital × Tasa`, `Total = Capital + Interés`, `Tasa` matches term (15→25%, 30→50%, 60/90→100%).
- **Payments — 61/61:** loan `Total Pagado` = Σ its `Pagos`. Lost records P-0001/P-0002 restored and reconcile.
- **Resumen — 34/34 clients:** count, Capital, Interés, Pagado, Saldo match the per-client sums.
- **Cash — to the peso:** Fondo 5,345,000 + Cobrado 8,328,500 − Prestado 13,090,000 = **583,500 = Efectivo disponible**.
- **Panel/Estadísticas** counts (activos 22, vencidos 7) match; no negative balances; `Saldo con Mora ≥ Saldo`.
- "Overpaid" PAGADO loans (e.g. L-0004 paid 349,000 vs total 225,000) = **mora collected**, `Saldo` correctly 0. Not errors.
- "Vencimiento ≠ Fecha+Plazo" rows = installment loans where `Vencimiento` is the last cuota date (by design). Not errors.

---

## Findings

### 🔴 N-1 · Installment schedules bill ~2× the loan total — 6 loans
`Cuotas` sum to double the loan's `Total a Pagar`:

| Loan | Total a Pagar | Cuota (actual) | Σ cuotas | Cuota (correct = Total/3) |
|---|---|---|---|---|
| L-0005 | 900,000 | 600,000 ×3 | 1,800,000 | 300,000 |
| L-0009 | 600,000 | 400,000 ×3 | 1,200,000 | 200,000 |
| L-0015 | 750,000 | 500,000 ×3 | 1,500,000 | 250,000 |
| L-0019 | 1,500,000 | 1,000,000 ×3 | 3,000,000 | 500,000 |
| L-0012 | 570,000 | 336,666.67 ×3 | 1,010,000 | 190,000 |
| L-0016 | 600,000 | 360,000 ×3 | 1,080,000 | 200,000 |

**Root cause (traced):** `cuotasForLoan_` (`LoanManagerV2.gs:1362`) computes `base = round2_(total / n)` — correct. So the **caller passed a doubled `total`**, or the loan's `Total a Pagar` was later edited down without regenerating cuotas. Creation sites: `LoanManagerV2.gs:6089` and `:8945` pass `totalPlan`/`total` + `ti.cuotas` (from `termForAmount_` → `bigTierCuotas_`, `Validaciones.gs:445`). Note these are **term-30** loans (rate 0.5) yet have 3 cuotas — the amount-based "tramo grande" splitter, not `installmentCount_` (which returns 3 only for term 90/3, `:3365`). **Impact:** a borrower following the installment plan / the "Estudio de Recordatorios" row / the per-cuota reminder pays nearly double. Whole-loan reads (`loanWaStatus_`, pre-due email) still show the correct total.

### 🟠 N-2 · Cuota payments not synced to the schedule — 2 loans
L-0012 (loan paid 160,000; cuota 1 shows 30,000) and L-0016 (loan paid 120,000; cuotas show 0). The `Monto Pagado`/`Saldo Cuota` cells are static, not the live `SUMIFS`/`VLOOKUP` formulas `cuotasForLoan_` installs (`:1383-1392`) — likely overwritten as static during the day-based migration. **Re-running `cuotasForLoan_` reinstalls the formulas and fixes this automatically** (so N-2 is repaired together with N-1).

### 🟡 N-3 · Orphan cuotas
`L-0051-1` (225,000) and `L-0053-1` (375,000) reference loans **L-0051 / L-0053 absent from Prestatarios**. Single cuota each. Loans were likely deleted and the cuotas left behind.

### 🟡 N-4 · Corrupted / mixed-type DNI storage
Raw `DNI` column holds inconsistent types: clean number (C-0002 `20534660`), text-with-dots (C-0017 `"39.594.460"`), and **decimal-mangled** — **C-0014 raw DNI = `25.390169`** (DNI 25,390,169 entered as "25.390.169" → parsed as 25.39, digits lost; recoverable only from `DNI (norm)`). Anything reading the raw DNI shows garbage. Concrete case of code-audit **W-4**.

### 🟡 N-5 · C-0002 email columns disagree
`Correo` = `cristinamabelavila68@gmail.com` vs `Correo (norm)` = `cristinamabelavila51@gmail.com` — two different addresses. Mail may misroute.

### 🟡 N-6 · Outstanding capital exceeds the lending fund — verify
Non-settled capital ≈ **8,300,000** vs `Fondo total para prestar` **5,345,000**. Cash is positive and reconciles (recycled collections), so not insolvency, but it appears to run past **V-09**'s intent. Confirm V-09's threshold/definition. Ties to code-audit **H-4** and your prior fix #12.

---

## Change / Repair Plan

Ordered by risk. Each item = **repair the data** (live sheet) and, where relevant, **fix the code** (prevent recurrence). Re-run the Python audit after each phase to confirm.

### Phase A — Installment schedules (N-1 + N-2) — highest priority
**A0. Decision gate (needs owner input):** confirm these 6 term-30 loans are *meant* to be 3-installment (the "tramo grande" feature) vs single-payment. If single-payment, the repair is to **remove** their cuotas instead of correcting them.

**A1. Data repair (assuming 3-installment is intended):** for each of L-0005, L-0009, L-0012, L-0015, L-0016, L-0019, rewrite the schedule so `Σ Monto Cuota = Total a Pagar`, **preserving the existing due dates** (Oct/Nov/Dec 13) and reinstalling the live payment formulas. Reuse `cuotasForLoan_` with the current total and the existing dates — proposed one-off:
```javascript
function repararCuotasDobles_() {               // correr una vez en el editor
  const ss = getSS_(), cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS);
  const IDS = ['L-0005','L-0009','L-0012','L-0015','L-0016','L-0019'];
  IDS.forEach(function (id) {
    const loan = findLoanById_(id); if (!loan) return;
    const dates = (cuotasDeLoan_(id) || []).map(function (c) { return c.due; })
                   .filter(function (d) { return d instanceof Date; });
    const n = dates.length || installmentCount_(loan);
    cuotasForLoan_(cs, id, loan.loanDate, round2_(loan.totalDue), n,
                   null, dates.length ? { dates: dates } : undefined);   // Σ = Total; reinstala fórmulas de pago
  });
}
```
(Alternative: `regenerarCuotasPrestamo_(id, 3)` — but that *restructures*: freezes debt to today and moves the first cuota to next week, changing the schedule. Use only if you want fresh dates.)

**A2. Code fix (prevent recurrence):** trace the `totalPlan`/`total` argument at the creation sites (`LoanManagerV2.gs:6089`, `:8945`) and `bigTierCuotas_` (`Validaciones.gs:445`) to find where the total is doubled (e.g. `amount*(1+rate)` applied twice, or `total` vs `per-cuota` confusion). Add an assertion/guard: after writing cuotas, verify `Σ Monto Cuota ≈ Total a Pagar` and log/throw otherwise. Also add a safeguard so editing a loan's amount **invalidates/regenerates** its cuotas.

**A3. Verify:** re-export and re-run `audit_numbers.py`; the "Σ Monto Cuota = Total a Pagar" and "Saldo Cuota = Monto − Pagado" checks must be 0 discrepancies, and L-0012/L-0016 cuota `Monto Pagado` must now reflect real payments.

### Phase B — Orphan cuotas (N-3)
Confirm L-0051/L-0053 are truly deleted (not renamed). If deleted: `deleteCuotasRowsForLoan_(cs,'L-0051')` and `'L-0053'`. If they should exist: restore the loan rows. Add a periodic integrity check (orphan Cuotas/Pagos → log to the `Errores` sheet). The existing `repararCuotas_` (`:1454`) only dedups repeated IDs, so orphans need this separate sweep.

### Phase C — Client data (N-4, N-5)
- **N-4:** format the `DNI` column as **plain text** and repair C-0014's raw cell from `DNI (norm)` (25,390,169). Sweep all clients for raw-vs-norm digit mismatches and repair from norm. Prevention: on intake, store DNI as text and strip separators before write (dovetails with code-audit **C-1 `safeCell_`**).
- **N-5:** determine C-0002's correct address and sync `Correo` and `Correo (norm)`. Add a validation that flags rows where `normEmail_(Correo) ≠ Correo (norm)`.

### Phase D — Fund exposure (N-6)
Review `V-09`: define whether the cap is on **cumulative** or **current outstanding** capital and against which fund figure; confirm current state (8.3M outstanding vs 5.345M fund) is intended (recycled capital) or a breach. Coordinate with code-audit **H-4** (caps evaluated inside the approval lock).

### Phase E — Regression guard
Keep `audit_numbers.py` + `audit_drill.py` (currently in the session tmp) as a repeatable **monthly reconciliation**: export the sheet, run both, review discrepancies. Consider porting the core checks into an in-sheet `auditarNumeros_()` that writes findings to the `Errores encontrados` sheet.

---

## Verification checklist (after repairs)
- [ ] `audit_numbers.py`: 0 discrepancies for Σ-cuotas, Saldo Cuota, orphan cuotas.
- [ ] L-0012 / L-0016 cuota `Monto Pagado` matches loan `Total Pagado`.
- [ ] No raw-vs-norm DNI/email mismatches in Clientes.
- [ ] Cash still reconciles (Fondo + Cobrado − Prestado = Efectivo disponible).
- [ ] Pricing, payment, and Resumen checks remain 0 discrepancies (no regressions).
