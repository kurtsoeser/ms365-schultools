# Stammdaten-Speicher: v2 kanonisch (Epic A)

## Kurz

| Schlüssel | Rolle |
|-----------|--------|
| `ms365-schooltool-data-v2` | **Kanonisch** – Schuljahre, Core, Setup |
| `ms365-tenant-settings-v1` | **Legacy** – wird nicht mehr automatisch geschrieben |

`ms365TenantSettingsLoad()` liest **zuerst v2**, sonst einmalige Migration aus v1.

## Schreiben (Stand Epic A)

- `ms365TenantSettingsSave()` schreibt nur noch über **`ms365AppDataV2.setCoreFromTenantSettings`**.
- `writeTenantSettingsV1Mirror()` ist **hart deaktiviert** (`V1_MIRROR_WRITE_ENABLED = false`).

## Browser-Backup

`legacyMirror.writeEnabled: false` im Export-Metadatenblock.

## Import-Adapter

Siehe `docs/` und `stammdaten-import-pipeline.js`.

## UI-Naming

- **Schulregister** = Werkzeug / Seite / Modi (Pflegen · Einspielen · Synchron)
- **Stammdaten** = Inhalt (Fächer, Lehrer, Klassen, …)

## Für Tool-Entwickler

Kein direktes Schreiben auf `ms365-tenant-settings-v1`. Nutzen: `ms365TenantSettingsLoad()` / `ms365AppDataV2`.
