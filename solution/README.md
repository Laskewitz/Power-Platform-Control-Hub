# Power Platform Control Hub — Solution

This folder contains the **Power Platform Control Hub** Dataverse solution, which provides supporting components for the Code App.

## Contents

| Path | Description |
|------|-------------|
| `PowerPlatformControlHub_1_0_0_0.zip` | Packaged unmanaged solution (import-ready) |
| `src/` | Unpacked solution source (managed by `pac solution unpack/pack`) |

## Solution Components

### Entities

#### `ppa_resourcetombstone`
Tracks resources that have been deleted or soft-deleted across Power Platform environments. Used by the Code App to maintain an accurate inventory and surface tombstoned resources in governance workflows.

## Working with the Solution

### Re-pack after changes

```powershell
pac solution pack --zipfile .\PowerPlatformControlHub_1_0_0_0.zip --folder .\src --packagetype Unmanaged
```

### Unpack a new zip

```powershell
pac solution unpack --zipfile .\PowerPlatformControlHub_1_0_0_0.zip --folder .\src --packagetype Unmanaged
```

### Import into an environment

```powershell
pac solution import --path .\PowerPlatformControlHub_1_0_0_0.zip --environment <environment-id>
```
