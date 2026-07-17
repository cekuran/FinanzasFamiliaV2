# Finanzas Familia

Personal/family finance webapp. Google Apps Script backend (`code.js`) + single-file frontend (`Index.html`) + Google Sheets as the database. Deployed as a webapp, multi-user isolated by `owner_email`.

## Deploy

```bash
npx clasp push
```

No local loop — refresh the deployed webapp URL to verify changes. Timezone: `Europe/Madrid` (`appsscript.json`).

## Files

- `code.js` — backend: `SCHEMA`, CRUD, validation. Functions exposed via `google.script.run`.
- `Index.html` — frontend: HTML + CSS + vanilla JS in one file. UI in Spanish.
- `appsscript.json` — webapp config (`executeAs`: USER_DEPLOYING, access: ANYONE).
- `.clasp.json` — clasp scriptId.

## Data model

- One sheet per entity; rows = records, columns = fields. Each row carries an `owner_email` for multi-tenant filtering.
- **Cuentas** → **Subcuentas**: via `parent_id`. One level of nesting only.
- **Transacciones**: `cuenta_id` + optional `subcuenta_id`. Account saldo = `saldo_inicial` + sum(tx delta); parent saldo includes its subcuentas' saldos.
- `SCHEMA` (top of `code.js`) is the source of truth. `migrarEsquema()` adds missing columns idempotently on bootstrap; new columns append at the sheet's end and the next `upsertFila` realigns.

## Conventions

- One entry point for data: `Estado.bootstrap = await call('bootstrap')`, then render.
- Mutation flow: `await call(...)` → assign to `Estado.bootstrap.*` → call a `refrescarX()` helper that re-renders affected views, repopulates filter dropdowns, and clears stale filter state. Direct `renderX()` calls bypass this — prefer the helper.
- Server-side validation; frontend shows `toast('Error: ' + e.message)` on rejection.
- Design tokens: colors (`--primary*`, `--error*`), spacing (`--s-1`…`--s-9`), radii (`--r-sm/md/lg/xl`). Mobile-first: bottom-nav <900px, sidebar ≥900px.
- Modals use class `.sheet` (bottom sheet on mobile, right rail on desktop). Form action bars use `position: sticky; bottom: 0` with negative margin to clear the sheet's safe-area padding.

## Frontend map

One route per view, one render function per route. Mutations route through helpers that keep filter state consistent.

**Routes** (`ROUTES` const): `dashboard`, `cuentas`, `movimientos`, `presupuesto`, `recurrentes`, `conciliar`, `categorias`. `go(route)` is the only nav entry; it toggles `.view.is-active` and calls the matching render.

**Estado** branches:
- `bootstrap` — snapshot from `call('bootstrap')`: `cuentas`, `categorias`, `transacciones`, `recurrentes`, `presupuestos`, `resumen`, `sesion`, `version`.
- `filtros` — `{ cuenta_id, subcuenta_id, tipo, desde, hasta, q }`. Bound to `#filtro*` inputs via `FILTRO_KEY`.
- `pres` — `{ anio, mes }` for presupuesto month nav.
- `txTipo` — current seg-state in tx sheet (`gasto`|`ingreso`|`transferencia`).
- `editing`, `recPlantillaId` — ephemeral.

**Renders**: `renderDashboard`, `renderCuentas`, `renderMovimientos`, `renderRecurrentes`, `renderConciliar`, `renderCategorias`, `renderPresupuesto`.

**Bootstrap utilities**: `call(fn, ...args)` (google.script.run promise wrapper), `normalizarBootstrap`, `recargar()` (full reload), `toast(msg)`.

**Sheets/modals**: `openSheet(id)` / `closeAllSheets()` open/close `.sheet`. `openTx`, `openCuenta`, `openCat`, `openRec` fill the form then open the sheet. `setTipoTx` and `actualizarSubcuentasTx` mutate visibility inside the tx sheet.

**Filters & refresh helpers**: `poblarSelect`, `poblarFiltros`, `actualizarFiltroSubcuenta`. After any cuenta mutation call `refrescarCuentas()` (also `poblarConCuenta` for the conciliar dropdown) — it runs `limpiarFiltrosInvalidos` then re-renders dependent views.

**DnD**: `wireDragReorder` on `.subcta-row` elements inside `renderCuentas`; calls `reordenarSubcuentas` server-side, then `refrescarCuentas()`.

## Gotchas

- No dev server — `clasp push` then refresh the deployed URL.
- HtmlService strips/limits certain HTML/CSS; avoid heavy client-side frameworks.
- Every row is tenant-scoped via `owner_email`; never assume cross-user visibility.

