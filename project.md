# Finanzas Familia — Project doc

Webapp de finanzas personales/familiares. Cada usuario ve solo sus datos, en cualquier dispositivo, sin instalar nada.

## Stack

- **Backend**: Google Apps Script (V8 runtime). Funciones expuestas vía `google.script.run`.
- **DB**: Google Sheets, una hoja por entidad (`Cuentas`, `Transacciones`, `Categorias`, `Presupuestos`, `Recurrentes`, `Conciliacion`).
- **Frontend**: Una sola página servida por `HtmlService`. HTML + CSS + JavaScript vanilla en un único `Index.html`. Sin build step.
- **Multi-tenant**: columna `owner_email` en cada fila. El bootstrap filtra por el email del usuario autenticado.
- **Timezone**: `Europe/Madrid` (`appsscript.json`).
- **Auth scopes** (`appsscript.json`):
  - `https://www.googleapis.com/auth/spreadsheets` — leer/escribir las hojas.
  - `https://www.googleapis.com/auth/userinfo.email` — identificar al usuario (`owner_email`).

## Arquitectura

```
┌─────────────────────────────────────────────────┐
│  Navegador                                      │
│  ┌───────────────────────────────────────────┐  │
│  │ Index.html (UI + JS vanilla)              │  │
│  │  - Estado (bootstrap, filtros, pres, …)  │  │
│  │  - Helpers: refrescarCuentas(), call(),  │  │
│  │    poblarFiltros(), wireDragReorder()    │  │
│  └──────────────────┬────────────────────────┘  │
└─────────────────────┼───────────────────────────┘
                      │ google.script.run (async)
                      ▼
┌─────────────────────────────────────────────────┐
│  Google Apps Script (code.js)                   │
│  - SCHEMA, migrarEsquema()                      │
│  - CRUD: guardarX / eliminarX / obtenerX        │
│  - bootstrap() arma el snapshot que consume     │
│    el frontend                                  │
└──────────────────┬──────────────────────────────┘
                   │ SheetsApp
                   ▼
┌─────────────────────────────────────────────────┐
│  Google Sheets (1 workbook compartido)          │
│  Hojas: Cuentas, Transacciones, …               │
│  Cada fila: id, owner_email, …campos entidad    │
└─────────────────────────────────────────────────┘
```

El ciclo de mutación estándar:
1. Frontend llama a una función GAS (`guardarX`, `eliminarX`, …).
2. GAS valida y escribe en la hoja.
3. Frontend recibe el array actualizado, lo asigna a `Estado.bootstrap.*`, y llama a un `refrescarX()` que re-renderiza las vistas afectadas y limpia filtros huérfanos.

## Prerrequisitos

- Node.js 14+ y `npx clasp` (`npm i -g @google/clasp`).
- Cuenta de Google con acceso a Apps Script y al spreadsheet del proyecto.
- `clasp login` con esa cuenta.

## Setup inicial

1. Crear el spreadsheet destino (puede estar vacío; `migrarEsquema()` crea las hojas en el primer arranque).
2. Crear un proyecto GAS nuevo (`clasp create`) o apuntar `scriptId` en `.clasp.json` al existente.
3. Compartir el spreadsheet con el usuario de deploy (`executeAs: USER_DEPLOYING`).
4. `npx clasp push` para subir el código.

## Deploy

```bash
npx clasp push            # sube cambios
npx clasp deploy          # publica nueva versión del webapp
```

Después de cada cambio, refrescar la URL del webapp para probarlo. No hay dev server local.

## Modelo de datos

### Cuentas y Subcuentas

- Una `Cuenta` top-level tiene `tipo` (`activo`/`pasivo`), `moneda`, `saldo_inicial`.
- Una `Subcuenta` cuelga de una cuenta top-level vía `parent_id`. **Un solo nivel**: una subcuenta no puede tener subcuentas.
- `saldo` de cuenta top-level = `saldo_inicial` + sum(delta tx sin subcuenta) + sum(saldo de cada subcuenta).
- `saldo` de subcuenta = sum(delta tx con esa `subcuenta_id`).

### Transacciones

- Vinculadas a `cuenta_id` (obligatoria) y opcionalmente a `subcuenta_id`.
- Tipos: `gasto`, `ingreso`, `transferencia`.
- `transferencia` ocupa dos movimientos con `cuenta_destino_id` y ratio de conversión.
- Las transferencias no admiten subcuenta.

### Recurrentes

- Plantillas que, al abrir la app, evalúan si deben generar movimientos nuevos y los insertan.
- Almacenan el JSON de la plantilla (incluye `subcuenta_id` si aplica).

### Presupuestos

- Categoría + importe esperado, reutilizable cada mes.

## Migración de esquema

`SCHEMA` (al inicio de `code.js`) es la fuente de verdad. `migrarEsquema()`:

- Compara columnas existentes vs `SCHEMA`.
- Añade las que faltan al final de la hoja.
- La primera escritura posterior (`upsertFila`) realinea filas al orden canónico.

Es idempotente — se puede llamar cada vez sin daño. Se ejecuta en `doGet`/bootstrap.

## Decisiones de diseño

- **Sheets como DB**: cero operaciones de DB; el spreadsheet mismo sirve como UI de auditoría. Trade-off: rendimiento y validación más limitadas a partir de cierto volumen.
- **Multi-tenant por columna** `owner_email` y no por spreadsheet separado: una sola fuente de verdad, un deploy, una URL de webapp.
- **Frontend vanilla, sin bundler**: HtmlService penaliza los frameworks por tamaño y CSP. Vanilla es suficiente para la escala esperada.
- **Subcuentas a un solo nivel**: evitar explosión combinatoria de permisos y visualizaciones. Si se necesitan jerarquías, probablemente toca cambiar de modelo.
- **Saldo canónico en conciliación**: `conciliar()` consulta `obtenerCuentas()` para evitar el bug de excluir txs de subcuentas marcado `conciliado`.
- **`refrescarCuentas()` central** post-mutación: garantiza que filtros, dropdowns y todas las vistas que dependen de cuentas se actualicen en un solo lugar.

## Limitaciones conocidas

- No hay suite de tests. Cambios se verifican manualmente abriendo el webapp.
- Sin build pipeline — todo se carga inline.
- HtmlService impone límites de tamaño (~50KB por HTML).
- Sin transacciones: si falla a mitad de una operación de varios pasos (p. ej. recurrente → múltiples tx), quedan filas parciales. Recurrentes generan tx una por una en backend, con reintento por tx fallida.
- Sin rate limiting en el frontend. GAS impone cuotas a nivel de proyecto.

## Roadmap

Ideas abiertas (no comprometidas):

- Importación desde extractos bancarios (CSV/OFX).
- Categorización automática (regex/ML sobre descripción).
- Vista multi-moneda consolidada con tipo de cambio histórico.
- Notificaciones cuando un presupuesto se acerca al 100%.
- Export a CSV/JSON para backups.
- Modo offline (PWA + IndexedDB).
