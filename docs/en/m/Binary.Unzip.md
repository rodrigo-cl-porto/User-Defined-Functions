# [`Binary.Unzip`](/src/m/Binary.Unzip.pq)

Extracts files from a ZIP archive and returns a table of entries with file names and decompressed content.

## Syntax
```fs
Binary.Unzip(
    ZIPFile as binary
) as table
```

## Parameters

- `ZIPFile` — A binary containing a ZIP archive (for example, the result of `File.Contents`).

## Return Value

A table with the following columns:
- `FileName` (text) — The entry name inside the ZIP.
- `Content` (binary or null) — The decompressed file content; `null` if decompression failed or entry unsupported.

## Example

```fs
Binary.Unzip(File.Contents("C:\Temp\archive.zip"))
```

This yields a table you can expand or transform. To read the content of the first file as text:

```fs
let
    Files = Binary.Unzip(File.Contents("C:\Temp\archive.zip")),
    FirstBinary = Files{0}[Content],
    FirstText = if FirstBinary <> null then Text.FromBinary(FirstBinary) else null
in
    FirstText
```

## **Credits**

- Author: Ignacio Barrau
- Source: [ExtractZIP.pq](https://github.com/ibarrau/PowerBi-code/blob/master/PowerQuery/ExtractZIP.pq)
