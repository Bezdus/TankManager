# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Build commands, the KOMPAS-3D dependency, the encoding rules, storage and the update process are in AGENTS.md (imported below). This file only adds what AGENTS.md leaves out.

@AGENTS.md

## Build notes
- `msbuild TankManager.sln /p:Configuration=Debug` (packages are `PackageReference`; run `msbuild /t:Restore` first on a clean checkout).
- There are no tests. To check a change, build it and run it against a live KOMPAS-3D instance with an `.a3d` assembly open.

## Encoding: current state
- Checked with `file`: `DrawingPreviewService.cs`, `FileLockDiagnostics.cs`, `MaterialAggregator.cs` and `UpdateService.cs` (all in `Core/Services/`) are Windows-1251 with no BOM.
- `ProductStorageService.cs` is now valid UTF-8, even though AGENTS.md still lists it as CP1251.
- Check a file's encoding before you edit it. Don't rely on the list above.

## Architecture (big picture)
- **Load pipeline** (`KompasService.LoadDocument` / `LoadActiveDocument`):
  1. Create a new `KompasContext`, which attaches to the running KOMPAS.
  2. Build `Product(topPart, context)`.
  3. `PartExtractor.ExtractParts` walks the `IPart7` tree recursively into a flat `List<PartModel>`.
  4. Parts with `ProductType.PurchasedPart` also go into `product.StandardParts`.
  5. `MaterialAggregator.AggregateMaterials` builds the material totals, then `product.NotifyAggregatesChanged()` runs.
- The `Product` keeps its `KompasContext`, so a live product holds COM references. Every call back into KOMPAS goes through that context, for example `ShowDetailInKompas` (which uses `PartFinder` to find the `IPart7` again from the `PartModel`, then `KompasCameraController`), plus drawing previews and laser cutting.
- **Two product sources:** products loaded live from KOMPAS, and products restored from JSON by `ProductStorageService`. Restored parts are `PartModelFromStorage`, a `PartModel` subclass that uses `new` to expose setters for restoring from the DTO. Such products have no KOMPAS context. When you add a persisted property to `PartModel`, update the storage DTO/mapping and `PartModelFromStorage` too.
- **Costing:**
  - `Core/Models/ManufacturingOperations.cs` defines `ManufacturingOperationBase` and its subclasses: laser cutting, bending, rolling and flanging.
  - `LaserCuttingService` computes laser cutting from a DXF file found by `DxfResolver`, which looks up a DXF by part marking in folders near the assembly.
  - Rates come from `PricingSettings`, stored as `pricing_settings.json` next to the exe and edited in `Views/PricingSettingsDialog`.
- **UI:**
  - `MainViewModel` (about 1,800 lines) owns nearly all state and commands. It creates `ProductStorageService` and `ExcelService` itself, and caches linked products in `_linkedProductsCache`.
  - `Views/*Panel.xaml` are UserControls with empty code-behind. They bind to the window's `DataContext`, often through `RelativeSource AncestorType=Window`.
  - `MainWindow.xaml` is large and holds much of the layout.
- **Excel export:** `ExcelService` uses ClosedXML.
- **Logging:** use `ILogger` / `FileLogger`.
