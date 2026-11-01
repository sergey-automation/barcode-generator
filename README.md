# Barcode Generator

Free EAN-13 barcode generator with Excel settings, batch generation, multiple sizes, frames, and PNG export.

The project is intended for simple Windows use: edit the settings in Excel, run the generator, and receive ready-to-use PNG barcode files grouped by size.

## What it does

- Generates EAN-13 barcodes as PNG files.
- Generates 1 to 50 sequential internal item numbers in one run.
- Default batch size: 10.
- Creates many barcode size variants automatically.
- Each size can be enabled or disabled.
- Each size can be generated:
  - without a frame;
  - with a frame;
  - both ways.
- Saves output into folders such as:
  - `out/40x20_mm/`
  - `out/55x30_mm/`
  - `out/60x30_mm/`
- Uses a normal Windows font for the digits below the barcode.
- The compiled Windows EXE does not require Python on the user's PC.

## End-user files

After the Windows EXE is built, the normal user needs:

- `barcode_generator.exe`
- `barcode_generator.xlsx`

The language packages also include a reference text explaining how the numbering scheme works.

## How to use

1. Put `barcode_generator.exe` and `barcode_generator.xlsx` in the same folder.
2. Open `barcode_generator.xlsx`.
3. On the `SETTINGS` sheet set:
   - internal prefix;
   - shift / supervisor code;
   - starting item number;
   - quantity (`1..50`, default `10`);
   - DPI;
   - overwrite flag (`0/1`);
   - digits-below-barcode flag (`0/1`).
4. On the `FORMATS` sheet:
   - `Active = 0` — disable the size;
   - `Active = 1` — enable the size;
   - `Mode = 0` — without frame only;
   - `Mode = 1` — with frame only;
   - `Mode = 2` — generate both.
5. Save the workbook.
6. Run `barcode_generator.exe`.
7. PNG files appear in the `out` folder next to the program.

Example:

`out/60x30_mm/2062700000016_60x30_frame.png`

## Default internal numbering example

The workbook uses this configurable internal structure:

`PP | BBB | NNNNNNN | C`

Where:

- `PP` — two-digit internal prefix.
- `BBB` — three-digit internal shift / supervisor code.
- `NNNNNNN` — seven-digit internal item number.
- `C` — EAN-13 check digit.

Default example:

`20 | 627 | 0000001 | C` → `2062700000016`

The next generated item increments only the seven-digit item number and recalculates the check digit.

### Shift / supervisor code

`BBB` is an internal company convention, not a GS1 rule.

Any unique value from `000` to `999` can be used.

A mnemonic option is to convert the first three Latin letters of a name to T9 keypad digits:

- `MAR → 627`
- `TAT → 828`
- `NAT → 628`

## EAN-13 check digit

For the first 12 digits `D1..D12`:

`S = (D1+D3+D5+D7+D9+D11) + 3×(D2+D4+D6+D8+D10+D12)`

Then:

`C = (10 - (S mod 10)) mod 10`

`C` becomes digit 13.

## Internal numbers and official external GTINs

The default workbook demonstrates an internal numbering scheme. It does **not** allocate or register globally unique GTINs.

If a product later receives an official external GTIN, do not treat it as a small edit of the internal number.

All of the first 12 digits may change:

- the internal prefix may disappear;
- the internal shift / supervisor code normally does not belong inside the external GTIN;
- the internal item number may be replaced by an externally managed item reference;
- digit 13 is recalculated from the new first 12 digits.

Shift, batch, and production traceability should then normally be stored separately, for example in `LOT / SERIE`, production date, an internal database, or additional marking.

For external commercial use, follow the GS1 rules applicable to your organisation and country.

## Barcode dimensions and compliance

The supplied dimensions are practical output presets for choosing a barcode that fits a label.

Generating an image at a particular width and height does **not by itself guarantee GS1 print compliance**. Final print quality depends on module width, quiet zones, printer resolution, scaling during layout/printing, contrast, substrate, and the scanner environment.

Do not stretch a generated PNG after generation if reliable scanning is important. Choose the required size preset instead.

## Russian and English packages

- `release_ru/` — Russian Excel workbook and Russian numbering reference.
- `release_en/` — English Excel workbook and English numbering reference.

Both use the same generator and the same workbook filename:

`barcode_generator.xlsx`

## Build the Windows EXE locally

Source files:

- `source/barcode_generator.py`
- `source/BUILD_EXE.bat`

On a Windows PC with Python 3 installed:

1. Open the `source` folder.
2. Run `BUILD_EXE.bat`.
3. The script installs/updates PyInstaller.
4. It builds `barcode_generator.exe`.
5. It copies the EXE into both language release folders.

The resulting EXE is standalone. Python and pip are not required on the target Windows computer.

## Automatic Windows build on GitHub

The repository contains:

`.github/workflows/build-windows.yml`

After uploading the repository to GitHub:

1. Open the `Actions` tab.
2. Select `Build Windows EXE`.
3. Click `Run workflow`.

GitHub builds the EXE on a Windows runner and publishes two downloadable artifacts:

- `Barcode-Generator-RU`
- `Barcode-Generator-EN`

Each artifact contains:

- `barcode_generator.exe`
- `barcode_generator.xlsx`
- the matching reference text.

## Repository layout

```text
barcode-generator/
├─ README.md
├─ .gitignore
├─ .github/
│  └─ workflows/
│     └─ build-windows.yml
├─ source/
│  ├─ barcode_generator.py
│  └─ BUILD_EXE.bat
├─ release_ru/
│  ├─ barcode_generator.xlsx
│  └─ REFERENCE_RU.txt
└─ release_en/
   ├─ barcode_generator.xlsx
   └─ REFERENCE_EN.txt
```

## License

Choose the repository license when creating the GitHub repository. No license file is bundled in this upload package.
