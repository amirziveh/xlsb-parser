# Security Policy

## Reporting a Vulnerability

This library parses binary XLSB files — a privileged attack surface for
untrusted inputs (file uploads, server-side ingestion, browser tabs). We
take parser bugs that can crash, hang, or exhaust memory seriously.

**Do NOT file a public GitHub issue for security vulnerabilities.**

Instead, email the maintainers at **security@example.com** with:

- A description of the issue
- A minimal reproducer (an `.xlsb` file or a synthetic byte sequence)
- The xlsb-parser version affected
- The runtime (Node version / browser) where you observed the issue

We will acknowledge receipt within 72 hours and aim to publish a fix within
30 days for high-severity issues.

## Known Limitations

- **Zip bombs**: `fflate`'s `unzipSync` will decompress the entire ZIP into
  memory. Use the `maxZipBytes` option to cap the decompressed size and
  refuse oversized inputs before they OOM your process.
- **No streaming unzip**: the entire `.xlsb` part is held in memory between
  phases. The streaming `openXlsb()` API only streams *rows*; the ZIP step
  itself is still buffered. A future release may replace `fflate` with a
  streaming zip reader.
- **Heuristic pivot-cache decoder**: when `parsePivotCaches: true`, the
  decoder uses byte-level type detection which can mis-parse field
  boundaries on non-trivial caches. The `fieldNames` list is trustworthy;
  `rows` should be treated as approximations.
