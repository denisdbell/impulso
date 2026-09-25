# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Google Apps Script project for **Impulso Crédito**, a loan-management system bound to a Google Sheets spreadsheet (locale es-AR, currency ARS). All UI text, sheet names, comments, and status values are in **Spanish** — keep new code and comments in Spanish to match.

There is no local build, lint, or test tooling (no clasp config). Code is deployed by pasting the `.gs` files into the spreadsheet's Apps Script editor (Extensions ▸ Apps Script); the web form/signing pages are published via "Deploy ▸ Web app". Verification is manual, in the spreadsheet.

## Files (one shared global namespace)

All four files live in the **same Apps Script project and share one global scope**. `Validaciones.gs` and `Reactivar.gs` deliberately reuse helpers defined in `LoanManagerV2.gs` (`getSS_`, `guard_`, `esc_`, `normDni_`, `normEmail_`, `sendBrandedEmail_`, `fmtMoney_`, etc.) — **never redefine a function that exists in another file**; duplicates break the whole project.

- **LoanManagerV2.gs** — core system: config and column maps, menu (`onOpen`), sheet setup, balance engine, contract PDF generation + email (`agreementHtml_`), payment receipts, lender sidebar, web app (`doGet`), virtual contract signing, daily reminder tasks, shared helpers, and the normalized Clientes model.
- **Validaciones.gs** — validation catalog (V-01…V-24), the "smart" intake web form (`intakeSmartHtml_`, `submitIntakeSmart`) that recognizes existing clients by Correo + DNI, client dedup/normalization, and the business-rule engine: term/rate tiers, per-client concentration cap, max simultaneous loans, repayment history and graduation limits (`validateApprovalV2_`).
- **Reactivar.gs** — post-migration module. Data was migrated to a **day-based** schema but left as static values; `reactivarFuncionalidad` idempotently reinstalls live formulas, data validations, conditional formatting, and protections, resolving columns **by header name** (`hkey_`/`headerIndex_`), not by index. Also contains the current day-based approval flow (`approveApplicantV2_`, `verifyApplicantV2_`, `activarDisparadores` installs the onEdit trigger) which **replaces the older month-based flow in LoanManagerV2.gs**.
- **SignatureData.gs** — data only: two embedded base64 data-URI signatures (`OWNER_SIGNATURE_DATAURI`, `COOWNER_SIGNATURE_DATAURI`) stamped on every contract. ~100k tokens on 13 lines — **do not read this file fully**; there is no logic in it.

## Architecture

### Normalized data model (3NF)
Client identity (Nombre/Correo/DNI/Teléfono) lives **only** in the `Clientes` sheet. `Prestatarios` (loans) references clients by `ID Cliente`; Nombre/DNI shown there are resolved from Clientes. `Pagos` references loans by `ID Préstamo`. `Cuotas` holds one row per installment.

### Column maps are the source of truth
Never hardcode column indices. LoanManagerV2 defines the maps: `PB` (Prestatarios), `PC` (Clientes), `PP` (Pagos), `CU` (Cuotas), plus `CFG.SHEETS` for sheet names and status constants `ST` (ACTIVO/VENCIDO/PAGADO/SALDADO), `CST` (cuota states), `SIGN` (contract signature states). Reactivar.gs instead resolves columns dynamically by normalized header name — follow whichever convention the file you're editing uses.

### Business rules (Validaciones.gs)
Terms are in **days** with rates derived from the term: 15→25%, 30→50%, 60/90→100% (`TERM_DAYS_RATES`); 90 days pays in 3 installments. Allowed term depends on amount tiers (`tier1Max_`/`tier2Max_`/`loanMax_`, configurable in the `Configuración` sheet). Approval runs concentration, loan-count, exposure, and graduation checks; an "Anular límites" override column exists on `Nuevos Prestatarios`.

### Entry points / flows
- `doGet(e)` in LoanManagerV2.gs routes the web app: default = smart intake form, `?page=saldo` = balance lookup, `?page=firmar&loan=…` = contract signing, `?page=clasico` = legacy form. Applications can be closed via a setting (`acceptingApplications_`).
- Intake: web form → `submitIntakeSmart` → appends to `Nuevos Prestatarios`.
- Approval: checking `Verificado?` on a row fires the installable onEdit → BCRA verification → `validateApprovalV2_` → creates/reuses client → moves loan to `Prestatarios` with signature state PENDIENTE, emails contract signing link. `Rechazar?` archives to `Rechazados`. Batch alternative: `procesarNuevosPrestatarios`.
- Setup functions run from the Apps Script editor or the "Gestor de Préstamos" menu: `setupConfirm` (rebuild sheets), `configurarValidaciones`, `reactivarFuncionalidad`, `activarDisparadores`.

### Caching gotcha
Reactivar.gs and Validaciones.gs memoize header maps and Clientes rows per execution (`_headerIndexCache`, `_clientesHdrCache`, `_clientesRowsCache`). If your code adds/moves columns or appends clients mid-execution, call the matching `invalidate…_()` helper.
