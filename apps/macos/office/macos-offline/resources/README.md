# Compiled macOS Office add-ins

This directory intentionally contains no fake Office files.

The production artifacts must be real macro-enabled Office packages with these fixed names:

- `VisualTeX.dotm`
- `VisualTeX.ppam`
- `addins.json`

Create the macro projects from the reviewed `.bas` sources, then package them with:

```bash
node scripts/package_macos_offline_addins.mjs \
  --word /absolute/path/VisualTeX.dotm \
  --powerpoint /absolute/path/VisualTeX.ppam
```

The packager rejects containers without a real `vbaProject.bin`, rejects missing VisualTeX module names, injects the reviewed `customUI14.xml`, preserves the fixed filenames, and writes SHA-256 metadata. The VisualTeX installer remains disabled until both validated artifacts exist.
