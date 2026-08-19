# AGENTS.md

Windows-only WPF desktop app (.NET Framework 4.8, C#) that integrates with
KOMPAS-3D (Russian CAD): reads `.a3d` assemblies via COM, builds a parts/BOM
view, exports to Excel, and syncs saved products to local/server storage.

## Build & verify
- Build: `msbuild TankManager.sln` or open in Visual Studio. Requires .NET
  Framework 4.8 Developer Pack (Windows). No test project, no CI, no linter.
- Entry point: `App.xaml.cs` -> `MainWindow` -> `MainViewModel` (constructed in
  the `MainWindow` ctor). MVVM with manual constructor injection (no DI container).

## KOMPAS-3D dependency (critical)
- The app does NOT start KOMPAS. `KompasContext` uses
  `Marshal.GetActiveObject("KOMPAS.Application.7")` and returns `null` when
  KOMPAS-3D isn't running, so load/link/save features silently fail without it.
- Interop DLLs are referenced from `..\Common\Kompas*.dll` (a SIBLING directory
  outside this repo), not the copy in `KompasApiDll\`. Building requires
  `..\Common` to exist.
- COM objects are manually released (`Marshal.ReleaseComObject`); be careful when
  editing `KompasContext`/`ComObjectManager`.

## File encodings (gotcha)
- Source files have MIXED encodings: some UTF-8 (with BOM), some Windows-1251
  (no BOM). Preserve the existing encoding per file when editing; re-saving a
  CP1251 file as UTF-8 corrupts its Cyrillic text. Known CP1251 files:
  `Core/Services/ProductStorageService.cs`, `Core/Services/UpdateService.cs`.

## Data & storage
- Saved products: JSON (`DataContractJsonSerializer`) under
  `AppDomain.CurrentDomain.BaseDirectory\products\<Name>_<Marking>\product.json`
  + `images\`. Server storage is an optional shared folder set at runtime,
  persisted in `storage_settings.json` next to the exe. All gitignored (under `bin\`).
- `FileLogger` writes `TankManager.log` in the working directory.

## Updates
- AutoUpdater.NET polls
  `https://raw.githubusercontent.com/Bezdus/TankManager/master/update.xml` on
  startup. Bumping the version means updating `update.xml` (also embedded as a
  Resource) and publishing a matching GitHub release.

## Conventions
- UI strings and most comments are in Russian; keep new UI text Russian.
- Commands use CommunityToolkit.Mvvm `RelayCommand` (`MainViewModel`); a local
  `RelayCommand<T>` in `MainWindow.xaml.cs` handles XAML Expander toggling.
