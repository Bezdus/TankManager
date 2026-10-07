# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows-only WPF desktop app (.NET Framework 4.8, C#) that integrates with
KOMPAS-3D (Russian CAD): reads `.a3d` assemblies via COM, builds a parts/BOM
view with materials and manufacturing-cost estimates, exports to Excel, and
syncs saved products to local/server storage.

## Build
- `msbuild TankManager.sln` (or Visual Studio). Requires the .NET Framework 4.8
  Developer Pack. NuGet packages use `PackageReference` (restore with `msbuild -restore`).
- No test project, no CI, no linter. Verify by building; runtime behaviour
  needs a running KOMPAS-3D instance.
- Old-style (non-SDK) `.csproj`: every new `.cs` file must be added as
  `<Compile Include=...>` and every new `.xaml` as `<Page Include=...>` in
  `TankManager.csproj`, or it won't be built.

## KOMPAS-3D dependency (critical)
- The app does NOT start KOMPAS. `KompasContext` attaches via
  `Marshal.GetActiveObject("KOMPAS.Application.7")`; if KOMPAS isn't running,
  load/link/preview/laser-cutting features silently no-op.
- Interop DLLs are referenced from `..\Common\Kompas*.dll` (a SIBLING directory
  outside this repo). Building requires `..\Common` to exist.
- COM objects are released manually (`ComObjectManager`,
  `Marshal.ReleaseComObject`); the `IApplication` from `GetActiveObject` is
  intentionally never released. Be careful editing `KompasContext` / `PartExtractor`.

## Architecture
- Startup: `App.xaml.cs` (update check, global exception handlers) →
  `MainWindow` → `MainViewModel` (constructed in the `MainWindow` ctor). MVVM
  with manual constructor injection (no DI container). Nearly all UI lives in
  `MainWindow.xaml`; `Views\PricingSettingsDialog` is the only separate view.
- Load pipeline (`KompasService.LoadDocument` / `LoadActiveDocument`, called
  from the VM via `Task.Run`): `KompasContext` → `new Product(topPart, context)`
  → `PartExtractor` walks the assembly into `PartModel`s → `MaterialAggregator`
  → `AttachLaserCutting`.
- Laser cutting: `DxfResolver` finds candidate DXF folders near the assembly
  file and matches DXFs by part marking (sheet-material parts only);
  `LaserCuttingService` measures cut/engraving length through KOMPAS and adds a
  `LaserCuttingOperation` to `PartModel.Operations`.
- Costing: `ManufacturingOperationBase` subclasses (`LaserCutting`, `Bending`,
  `Rolling`, `Flanging` in `Core\Models\ManufacturingOperations.cs`) implement
  `CalculateCost(PricingSettings)`. `PricingSettings` are shared: the master copy is
  `<server folder>\_pricing_settings.json`, `pricing_settings.json` next to the exe is the
  offline cache (`PricingSettings.SyncWithServer`; equal mtimes = cache in sync). They are
  edited in `PricingSettingsDialog`, which is read-only in viewer mode. Stored products
  are re-costed with the current prices on open (`LoadAndLinkProduct`).
  Laser cutting price (руб/м; cut length is in mm) depends on sheet thickness: `PricingSettings.LaserCuttingPricing`
  (exact thickness → price, otherwise `LaserCuttingPricePerMeter`; the old руб/мм value
  `LaserCuttingPricePerMm` is converted ×1000 on load); the thickness is parsed from the
  material string (`PartModel.ParseSheetThickness`) and set on the operation in `RecalculateAllCosts`
  and `OperationsEditorDialog` (not persisted).
  Adding an operation type requires updating the enum, the subclass, and the
  `ToOperationDto`/`FromOperationDto` mapping in `ProductStorageService`.
  Subclasses implement only `CalculateUnitCost`; the base `CalculateCost` applies the manual edits
  made in `Views\OperationsEditorDialog` (engineer or technologist, applied to all instances of the part):
  `Origin` (Kompas/Manual), `Quantity`, `IsExcluded` (KOMPAS operations are excluded, never deleted),
  `HasManualCost`/`ManualCost`, `IsEdited`. `CustomOperation` = user-defined name + price per unit.
  Edits are saved immediately, separately from `product.json`, to `operations.json` in the product folder
  (`ProductStorageService.OperationEdits.cs`): one entry per part key `Name|Marking|FilePath` holding all
  operations of the part (empty list = edits reset), merged per part by `ModifiedUtcTicks` between local and
  server (`SyncOperationEditsFolder`; `CopyProductFolder` skips this file). `Load` and KOMPAS loads
  (`RestoreFromSaved`) apply it with `OperationEditsMerger.ReplaceEdits` (replaces edits stored in
  `product.json`; KOMPAS operations matched by type + ordinal).
- Storage (`ProductStorageService`): products are saved as JSON DTOs
  (`ProductDto`/`PartModelDto`/`OperationDto`, `DataContractJsonSerializer`)
  under `<exe dir>\products\<Name>_<Marking>\product.json` + `images\`. An
  optional shared server folder (set at runtime, stored in
  `storage_settings.json` next to the exe) is synced with the local copy.
  Writes go through `AtomicFile` (temp file + replace) and all mutating storage
  operations take `_ioLock`. Deleting "everywhere" records a tombstone
  (`_deleted.json` on the server, `_pending_deleted.json` locally while the server is
  unreachable) that `SyncFromServer` applies, so deleted products don't come back.
  Save/sync errors are not swallowed: local failures throw, server failures land in
  `ProductStorageService.LastServerError`.
  Without its own `storage_settings.json` the app reads `storage_settings.default.json`
  (shipped in the release zip, holds the shared server folder path).
- App modes (`Core\Services\AppMode.cs`): engineer (KOMPAS load/save/delete), viewer
  (procurement etc. without KOMPAS: read-only, products come from the server) and technologist
  (viewer + editing operations: `IsViewer` and `IsTechnologist` are both true, `CanEditOperations`;
  only `operations.json` is written to the server, pricing stays read-only). Decided once
  at startup in `MainViewModel` ctor: setting `Mode` in `storage_settings.json`
  (Auto/Engineer/Viewer/Technologist, switched via the mode chip in the top toolbar (`ModeButton` → `ModePopup`; the current mode is also in the window title), applies after restart); Auto =
  viewer when `KOMPAS.Application.7` isn't registered. Viewer never writes to the server:
  `ProductStorageService` checks `DownloadOnly` (sync phase 2, tombstones, image sync,
  Save/Delete). KOMPAS-only UI is bound to `IsEngineerMode`. Sync also runs in the
  background on startup and when the products panel opens (`RunServerSyncAsync`).
- `FileLogger` writes `%AppData%\TankManager\TankManager.log`. Use `ILogger`, not
  `Debug.WriteLine` (no output in Release).
- COM lifetime: `Product` owns its `KompasContext` and disposes it (`Product.Dispose`);
  `MainViewModel` disposes products that are replaced and not cached. All KOMPAS calls in
  `KompasService` are serialized by `_kompasLock`. Documents opened hidden via
  `Documents.Open(..., false, ...)` must be closed (see `KompasContext.CloseIfHidden`).

## File encodings
- All `.cs`/`.xaml` files are UTF-8 with BOM, CRLF (see `.editorconfig`). Never re-save a
  file in another encoding: Cyrillic text gets destroyed (this already happened once to
  `ProductStorageService.cs` and was restored from git history).

## Releases / updates
- AutoUpdater.NET polls
  `https://raw.githubusercontent.com/Bezdus/TankManager/master/update.xml` on
  startup. A version bump means updating `Properties/AssemblyInfo.cs`
  (`AssemblyVersion`/`AssemblyFileVersion`), `update.xml` (also embedded as a
  Resource), and publishing a matching GitHub release zip
  (`TankManager-vX.Y.Z.zip`).

## Known limitations
- Rolling ("вальцовка") cost = part mass × `RollingPricePerKg`; the mass is set by
  `MainViewModel.RecalculateAllCosts` (`RollingOperation.PartMass`, not persisted). The old
  `RollingPricePerMm` setting was replaced, so the rolling price must be re-entered once.
- `FilePath`/`DxfFilePath` from `product.json` may point to network shares; only preview PNG paths
  are restricted to the local products folder. Treat a writable server folder as trusted.
- `update.xml` has no `<checksum>`: add SHA-256 of the release zip when publishing a release.

## Conventions
- UI strings, comments, and commit messages are in Russian; keep new UI text Russian.
- Commands use CommunityToolkit.Mvvm `RelayCommand` (`MainViewModel`); a local
  `RelayCommand<T>` in `MainWindow.xaml.cs` handles XAML Expander toggling.


## KOMPAS SDK documentation

The KOMPAS-3D SDK COM API v7 docs (~1300 Markdown files, in Russian) live in `ai_docs/` at the repo root. They are **generated** by `python tools/build_ai_docs.py` from the raw HTML help of SDK v23 in `KOMPAS_SDK_ru-RU/` (14 000 pages; its table of contents `js/hmcontent.js` drives the structure). Both folders are **gitignored**, so either may be absent on a fresh clone — if `ai_docs/` is missing but the raw help is there, run the generator; if both are missing, say so rather than guessing API names. Don't hand-edit generated files: fix the generator and rerun it. Only `README.md` and `guides/csharp_setup.md`, `events_v7.md`, `automation.md` are hand-written (the generator leaves them alone; the C# snippets in them were compiled against `KompasMCP/libs/`). Consult the docs before writing any KOMPAS API code:

| Path | Contents |
|------|----------|
| `ai_docs/README.md` | Where the files come from, how to search, ProgIDs, document types, the main object chains. Read first (5 KB). |
| `ai_docs/SUMMARY.md` | Index (~1650 lines, 120 KB), interfaces grouped by help section (*Документ 3D / Компоненты* …); sections *Интерфейсы COM v7*, *Интерфейсы событий*, *Перечисления и константы*, *Руководства*. Keyword search starts here. |
| `ai_docs/api/` | 785 files, one per v7 interface. Each holds the hierarchy, notes, and **every property/method inline** as `### Name - description` under `## Свойства` / `## Методы`, with Automation + COM syntax. |
| `ai_docs/enums/` | 478 enumerations and constants — all v7 enums plus the v5-era constants pages, e.g. `obj3dtype.md` (`ksObj3dTypeEnum` with the matching API7 interface per code), `kslengthunitsenum.md`. |
| `ai_docs/events/` | 39 event interfaces (API7 and API5 sections — `ksDocument3DNotify`, `ksDocumentFileNotify` live in the latter) plus who the event source is. |
| `ai_docs/guides/` | New in v23, compiling libraries, help for applications; hand-written `csharp_setup.md` and `events_v7.md` are the ones that matter here. |

Links to parts of the help that are not converted (API5, export functions, parameter structures) point straight into `KOMPAS_SDK_ru-RU/*.html`.

### How to search it

`SUMMARY.md` is too big to read whole — use `Grep`, never `Read` it end to end.

1. `Grep` a keyword in `ai_docs/SUMMARY.md` (Latin, e.g. `Circle`, `Extrusion`, or Russian, e.g. `Эскиз`) to find the interface and its `api/<name>.md` link. File names are the lowercase help page names (`ipart7.md`, `ksobj3dtypeenum` lives in `obj3dtype.md` — grep the heading, not the file name).
2. Open that `api/` file. Read the **Примечание** block near the top: it states required calls and how to obtain the object (e.g. "call `IDrawingObject::Update` after setting parameters"). Follow the **Иерархия** list to find inherited members — `ICircle` has no `Update`, it comes from `IDrawingObject`. Entries marked «Дополнительно:» are extra interfaces reached by a cast (QueryInterface): `IView` → `IDrawingContainer`, `IPart7` → `IModelContainer`.
3. To find one member without reading a big file, `Grep -n "^### Update"` in the interface file, then `Read` with `offset`/`limit`.
4. For an enum value, `Grep` the constant name across `ai_docs/enums/`. If unsure which enum a parameter takes, the member's section names the type and links to it.
5. There is no `samples/` folder. For usage examples look at `guides/csharp_setup.md` and the existing services (`SketchService`, `OperationService`, `GeometryService`).