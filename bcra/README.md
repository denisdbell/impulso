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

## Datos sensibles

`out/` (reportes y JSON con datos personales/financieros) y `cuils.txt` (CUILs reales)
están en `.gitignore` y **no se versionan**. Si querés compartir un reporte, copiá el
`.report.md` puntual fuera del repo.
