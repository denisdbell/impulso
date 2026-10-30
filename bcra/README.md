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
