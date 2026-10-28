# AutoSoft GridCraft fork

Updated: 2026-10-28. Based on upstream `11bb67386c9bb19b0717e21e98adc608162ae2e5`.

This independent fork is maintained at <https://github.com/technologiesautosoft-coder/gridcraft>.
The application identifies itself as `0.12.1-autosoft.1`. It is based on GridCraft by the ArtCraft
team and community, retains the MIT/Apache notices, and omits ArtCraft trademark graphics.

## What changes

XML Spreadsheet 2003 workbooks exported with an `.xls` extension open as worksheets instead of
raw XML lines. Detection uses the Workbook root's namespace, not the filename, so the same content
also imports from `.xml`. The existing BIFF8 `.xls` reader remains in place.

The XML importer keeps worksheet names, strings (including XML entities, UTF-8 and UTF-16), finite
numbers, booleans, errors and date/time values; sparse row/cell indices; merged cells; row heights
and column widths; hidden rows/columns; basic fonts, solid fills, borders, alignment and custom
number-format codes; and default/parent/row/column/cell style inheritance. Imported workbooks save
to a new `.xlsx` destination. Save never silently overwrites an `.xls` or `.xml` source.

macOS, Windows and Linux packages advertise `.xls` in their Open With registrations. macOS uses
the existing `com.microsoft.excel.xls` type and Alternate rank; it does not take over the default
application for existing workbooks.

## Current limits

XML formulas use saved results, with an import warning that they will not recalculate. A formula
without a saved result receives a visible unavailable-result marker. Live XML formula translation,
rich text, print/view options, names, charts, validation and conditional-format rules are not
implemented. Existing binary XLS formula support is unchanged. Non-default table offsets are
refused rather than shifted silently. Invalid numeric values, malformed XML, DTDs, unknown entities,
overlapping merges and out-of-bounds indices fail without changing the open document.

Import is capped at 32 MiB, 250,000 XML elements, depth 64, 4,096 styles, 256 sheets, and one million
expanded cells and dimension records (including merged areas). Styles cannot form inheritance cycles. External entities
are never resolved. Real operational workbooks stay outside the repository; tests use synthetic XML.

This fork checks updates only when requested in Help, using the AutoSoft fork's releases. It does
not prompt at startup or offer an upstream installer that would remove these changes. Until the
fork publishes release packages, update checks can report that no release is available.

## Development and local build

Acceptance builds use the same external `craft-fonts` input as upstream CI, pinned to
`b076417b5880bd4154bd22f00695b1d28c6bea83`. Set `CRAFT_FONTS_DIR` to that checkout's absolute path
and `CRAFT_FONTS_REQUIRED=1` before running the gate or building the application. The repository's
existing policy filters barred font families and limits the web build's font set. Copy the relevant
OFL notices into the app package alongside the other required notices. Without this input, the
bare egui test faces lack some glyphs used by the upstream interaction tests.

```sh
cargo test -p gridcraft-xlsx spreadsheetml
cargo test -p gridcraft-engine --test it spreadsheetml_import
cargo xtask ci
cargo build --release --locked -p gridcraft -p gridcraft-cli
```

On Apple silicon the upstream packaging script builds a locally signed installer with:

```sh
packaging/macos/package.sh --arch aarch64
```

Local ad-hoc signatures identify a local build; they are not Apple Developer ID notarization and
must not be represented as an official upstream release. Retain the previous app and settings
until the new application has opened the original XML report and Finder has offered it for `.xls`.

The XML reader is `crates/xlsx/src/spreadsheetml.rs`; format dispatch and safe import paths live in
`crates/engine/src/io.rs` and `crates/engine/src/cmd/file.rs`. OS registrations live in `packaging/`.
