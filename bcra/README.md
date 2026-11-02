# Descarga por lotes de la Central de Deudores (BCRA)

Script independiente (Python 3, solo biblioteca estándar) que consulta la API pública
de la Central de Deudores del BCRA para una lista de CUILs y guarda un archivo JSON por
persona. **No toca la planilla ni el proyecto Apps Script**: solo lee `cuils.txt` y
escribe en `out/`.

## Uso

1. Pegá los CUILs en [`cuils.txt`](cuils.txt), uno por línea (con o sin guiones).
2. Corré:

   ```sh
   python3 bcra/fetch_bcra.py
   ```

3. Resultado en `bcra/out/`:
   - `<cuil>.json` — datos crudos del BCRA por persona (Deudas + Históricas + Cheques
     Rechazados combinados en un solo archivo).
   - `_index.csv` — resumen de todos: CUIL, denominación, peor situación, entidades,
     cheques rechazados, estado HTTP.
   - `_errors.log` — CUILs que fallaron o líneas inválidas.

## Pensado para no ser bloqueado ni limitado

- Secuencial (un pedido por vez), con pausa de ~2–3 s entre pedidos (`--delay` / `--jitter`).
- Reintenta con backoff exponencial ante 429 / 5xx (respeta `Retry-After`).
- `User-Agent` descriptivo y `Accept: application/json`.
- Reanudable: si el archivo del CUIL ya existe, lo saltea (no vuelve a pegarle a la API).
  Usá `--force` para re-descargar.

## Opciones

```
python3 bcra/fetch_bcra.py --help

  --input cuils.txt       archivo de CUILs (uno por línea)
  --out out               carpeta de salida
  --endpoints all         'all' (Deudas + Historicas + Cheques) o 'historicas' (solo histórica)
  --delay 2.0 --jitter 1.0  espaciado entre pedidos
  --retries 4 --timeout 30
  --force                 re-descargar aunque exista el archivo
  --insecure              desactiva verificación TLS (recurso ante el certificado incompleto del BCRA)
```

> Si ves un error de certificado TLS contra `api.bcra.gob.ar`, probá agregando `--insecure`.
> El BCRA a veces publica una cadena de certificados incompleta; `--insecure` afecta solo
> el transporte, no los datos.

## Análisis / reporte (`analyze_bcra.py`)

Convierte los datos descargados en un **reporte de crédito en inglés** por persona:
reconstruye el historial de pagos, clasifica el comportamiento de pago en categorías,
calcula métricas estándar de riesgo y un puntaje 0–100 con calificación A–E, y **rankea**
a la persona contra el resto de los deudores ya descargados.

```sh
# Una persona → se imprime en pantalla y se guarda out/<cuil>.report.md
python3 bcra/analyze_bcra.py 20123456789   # <cuil> sin guiones (ejemplo)

# Tabla ordenada de todos los deudores descargados → out/_ranking.md
python3 bcra/analyze_bcra.py --all
```

- Si el `<cuil>.json` no existe todavía, se descarga primero (reutiliza la lógica cortés
  de `fetch_bcra.py`; una sola ruta HTTP con el mismo rate-limiting).
- **Perfil de pago** (según la serie de 24 meses): *Consistent — always current*,
  *Consistent — performing*, *On and off (volatile)*, *Recovering*, *Deteriorating*,
  *Chronically delinquent*, *No record*.
- **Puntaje ponderado por historia**: promedio de la situación mensual con decaimiento por
  recencia (la mora reciente pesa más y sigue bajando el puntaje aunque la deuda ya esté
  saldada). Por eso **difiere a propósito** del chequeo rápido (solo situación actual) de
  la planilla (`bcraScore_` en `LoanManagerV2.gs`).
- El “historial de pagos” se **infiere** de las bajas de saldo mes a mes: la Central de
  Deudores es una clasificación mensual, no un registro de transacciones.

Opciones: `--out`, `--no-save` (solo pantalla), `--no-fetch` (no descargar si falta),
y los mismos flags de red que `fetch_bcra.py` (`--delay/--jitter/--retries/--timeout/--insecure`).

## Historial Impulso + ranking combinado (`analyze_impulso.py` / `analyze_combined.py`)

Suma el **historial de préstamos propio de Impulso Crédito** al análisis: produce un
ranking de pago *interno* y, cruzando por CUIL, un **ranking combinado** (BCRA + Impulso).
Lee la planilla exportada a `.xlsx` (sin tocar la hoja ni el proyecto Apps Script).

1. Conseguí la planilla como `bcra/out/_impulso.xlsx` (ruta por defecto; queda en
   `out/`, ignorado). Dos caminos:

   **a) Automático — service account (sin navegador, sin Claude)** — `fetch_impulso.py`:

   ```sh
   pip install google-auth requests           # única dependencia (solo del descargador)
   python3 bcra/fetch_impulso.py              # exporta la hoja → out/_impulso.xlsx
   ```

   Setup por única vez (~5 min): en Google Cloud Console habilitá la **Google Drive API**,
   creá un **service account** con una **clave JSON**, guardala como `bcra/service_account.json`
   (ignorada), y **compartí la planilla como *Lector* con el `client_email` del JSON** (igual
   que compartís con una persona; si no, el export da 404). El id de la planilla se toma de
   `--file-id`, la variable `IMPULSO_SHEET_ID`, o el archivo ignorado `bcra/.impulso_file_id`
   (así el id real no vive en el código). El descargador es el **único** archivo que usa
   librerías de terceros; el análisis sigue siendo biblioteca estándar.

   **b) Manual (cero setup)** — en el Sheet, **Archivo ▸ Descargar ▸ Microsoft Excel (.xlsx)**
   y guardalo como `bcra/out/_impulso.xlsx`.

2. Corré. **`report.py` es el único comando que necesitás** — muestra los tres
   reportes (BCRA + Impulso + combinado) juntos:

   ```sh
   # Todos: imprime los 3 rankings y guarda los 3 .md
   python3 bcra/report.py --all

   # Una persona (CUIL, DNI o ID Cliente): veredicto combinado + reporte BCRA +
   # reporte Impulso, y guarda out/<id>.full.report.md
   python3 bcra/report.py 20123456789
   python3 bcra/report.py 12345678          # por DNI
   python3 bcra/report.py C-0007            # por ID Cliente
   ```

   El archivo BCRA de cada persona **se descarga solo si falta** (usa la ruta cortés
   y con límite de `fetch_bcra.py`): uno en el modo persona, y el de cada prestatario
   en `--all`. Para trabajar sin red usá `--no-fetch`; si el BCRA falla por su cadena
   de certificados TLS, agregá `--insecure`.

   O cada análisis por separado (lo que `report.py` orquesta por detrás):

   ```sh
   python3 bcra/analyze_impulso.py --all            # → out/_impulso_ranking.md
   python3 bcra/analyze_impulso.py 20123456789      # → out/<cuil>.impulso.report.md
   python3 bcra/analyze_combined.py                 # → out/_combined_ranking.md
   python3 bcra/analyze_combined.py 20123456789     # veredicto combinado de una persona
   ```

### Aplicación de escritorio (interfaz gráfica)

Para ver todo sin la terminal, abrí la app gráfica — **Tkinter** (viene con Python,
sin `pip install`), así que corre en cualquier lado:

```sh
python3 bcra/report_gui.py
```

Barra lateral con **Resumen** (tarjetas + gráfico de decisiones + mejores/peores),
tres pestañas de ranking (**Combinado / BCRA / Impulso**) buscables, ordenables y
coloreadas por riesgo (doble clic en una fila abre a la persona), y una vista de
**Persona** con medidores de puntaje (combinado / BCRA / Impulso), la tira de
préstamos de Impulso, la línea de tiempo de 24 meses del BCRA y las tablas de detalle.
Los números son exactamente los mismos que los de `report.py`. Los archivos BCRA
faltantes se bajan en segundo plano (se puede desactivar en *Options*); exportá
cualquier ranking o una ficha a Markdown desde el menú. Abrí otra planilla con
*File ▸ Open workbook…*.

The app adds two **anti-fraud** sections to the sidebar: **🚩 Fraude**
(cross-record signals, computed on load) and **🪪 DNI** (OCR verification, triggered
by a button). Both also appear as cards in the **Persona** view and export to
Markdown. The same reports are available from the command line.

> _Note: these two features (GUI text, labels and reports) are intentionally in
> **English**, matching the rest of the bcra tooling (`report_gui.py`,
> `analyze_*.py`). The literals that must match external data — the DNI card fields
> read by OCR, the workbook column/sheet names, and the `spa` OCR model — stay in
> Spanish._

### Fraud signals (`analyze_fraud.py`)

Mines the workbook itself for indications of suspicious identity/behavior, cross-
referencing sheets (Clientes, Dispositivos, Firmas, Prestatarios, Pagos, Nuevos
Prestatarios). **No OCR and no new dependencies** (it only reads the `.xlsx`).

```sh
python3 bcra/analyze_fraud.py                      # per-client ranking (to stdout)
python3 bcra/analyze_fraud.py --export out/_fraude.md
```

Signal catalog, with per-client weights (`HIGH=40 · MED=20 · LOW=10`) and a resulting
risk of **High** (≥40 or any HIGH signal) · **Review** (1–39) · **OK** (0):

- **Duplicate identity:** DNI or CUIL repeated across clients (HIGH), inconsistent
  CUIL↔DNI (AFIP check digit + embedded DNI), shared phone/email, invalid email,
  shared address, flagged as duplicate, blocked with an active loan (HIGH).
- **Device / network:** same **device fingerprint** (HIGH), same **IP** (≥3 clients),
  identical **GPS** (≥3), **automation/webdriver** (HIGH), **signing** IP shared
  across DNIs. Thresholds are configurable constants in the module.
- **Reference ring:** a reference phone equal to a client phone, or reused across
  several applications.
- **Third-party account / mule:** disbursement to a third party (`¿Cuenta propia?`=No);
  one account holder receiving for several borrowers (HIGH).
- **Consistency:** a loan's or payment's denormalized `Nombre/DNI` differing from the
  linked client.

These are **indicative leads for manual review, not verdicts** (several have
legitimate explanations: cohabitants, same neighborhood, low GPS precision).

### DNI verification by OCR (`fetch_dni.py` + `verify_dni.py`)

Downloads the DNI photos from the Drive folder, runs **local OCR (Tesseract)** on
them, and compares each document against the Clientes sheet, flagging anomalies
(`DNI mismatch`, `Name mismatch`, `Front and back mismatch`, `DNI not found in
Clientes`, `Document expired`, etc.). It is the Python port of `VerificacionDNI.gs`
(same folder, same fields), to keep the results consistent.

Install (one-time — local OCR engine, the images **never** leave the machine):

```sh
brew install tesseract tesseract-lang        # binary + Spanish data (spa)
pip install pytesseract pillow google-auth requests
```

The download uses the **same service account** as `fetch_impulso.py`. Share the DNI
folder (and its subfolders) **as Viewer** with the `client_email` from
`service_account.json`; otherwise the listing returns 404. The folder id comes from
`--folder-id`, the `DNI_FOLDER_ID` env var, a `.dni_folder_id` file, or the default
(the same id already present in `VerificacionDNI.gs`).

```sh
python3 bcra/fetch_dni.py        # → out/dni/<person>/…  (resumable, PII git-ignored)
python3 bcra/verify_dni.py       # OCR + comparison → Markdown report
python3 bcra/verify_dni.py --fetch --export out/_dni.md     # download and scan in one
```

OCR reading is **indicative**: it may contain errors and does not replace manual
verification of the document.

- **Puntaje Impulso** (0–100, misma escala A–E que el BCRA): comportamiento de pago real
  con nosotros por préstamo — *a tiempo / tarde / vencido / en plan / congelado* —
  ponderado por recencia, con piso de calificación según la peor situación vigente.
  Perfiles: *Excellent — always on time*, *Reliable*, *Occasional late payer*,
  *Chronically late*, *In a payment plan*, *Frozen (hardship)*, *Currently overdue*,
  *New — no closed loans yet*, *No record*.
- **Fecha de vencimiento efectiva** = `Fecha Préstamo + Plazo (días)` (regla del negocio).
  El `Vencimiento` guardado es una fórmula cuyo valor exportado está corrupto en unos pocos
  préstamos problemáticos (empujado meses al futuro); por eso se usa la fecha derivada y se
  marca con ⚠ la discrepancia en la ficha.
- **Combinado** = `60%·Impulso + 40%·BCRA` (el pago de primera mano pesa más por ser el
  mejor predictor; ajustable con `--impulso-weight`). La **decisión sale del puntaje
  combinado** (≥60 OK · ≥35 REVIEW · resto REJECT), se eleva a REVIEW si *alguna* fuente
  rechaza y se mantiene en REJECT si *ambas* lo hacen — un buen pagador nuestro nunca queda
  rechazado solo por el bureau. Quien figura en una sola fuente se arrastra con esa sola y
  se etiqueta *BCRA only* / *Impulso only*.
- **Sin dependencias**: `xlsx_mini.py` es un lector `.xlsx` de biblioteca estándar
  (`zipfile` + `xml`), validado celda por celda contra `openpyxl` sobre la planilla real.

## Datos sensibles

`out/` (reportes y JSON con datos personales/financieros, **incluida la planilla
`_impulso.xlsx`**), `cuils.txt` (CUILs reales) y cualquier `*.xlsx` están en `.gitignore`
y **no se versionan** — solo se versiona el código. Si querés compartir un reporte, copiá
el `.report.md` puntual fuera del repo.
