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
  `CalculateCost(PricingSettings)`. `PricingSettings` is persisted to
  `pricing_settings.json` next to the exe and edited in `PricingSettingsDialog`.
  Adding an operation type requires updating the enum, the subclass, and the
  `ToOperationDto`/`FromOperationDto` mapping in `ProductStorageService`.
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
- Rolling ("вальцовка") length is not read from KOMPAS yet (`KompasContext.ProcessModelObject`,
  TODO), so its cost is 0; `IsCostReliable` marks such operations and the UI warns about them.
- `FilePath`/`DxfFilePath` from `product.json` may point to network shares; only preview PNG paths
  are restricted to the local products folder. Treat a writable server folder as trusted.
- `update.xml` has no `<checksum>`: add SHA-256 of the release zip when publishing a release.

## Conventions
- UI strings, comments, and commit messages are in Russian; keep new UI text Russian.
- Commands use CommunityToolkit.Mvvm `RelayCommand` (`MainViewModel`); a local
  `RelayCommand<T>` in `MainWindow.xaml.cs` handles XAML Expander toggling.
