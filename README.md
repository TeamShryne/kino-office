# Kino Office

Optional Kino runtime library for creating and editing Word, Excel, and PowerPoint files.

This repository builds an Android library (AAR), converts its classes to Android DEX, packages them as a Kino runtime library, and publishes a GitHub Release. Kino downloads the matching ABI package on demand.

## Runtime package format

- `classes.jar` containing `classes.dex`
- `library.json`

No native library is required.
