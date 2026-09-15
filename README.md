# Base64 Tools

[Get the complete Base64 bundle](./BASE64_TOOLS_COMPLETE.base64.txt)

One complete, no-comment bundle containing all **58 VBA modules** from the September 15, 2026 SLIC-preserving release. The earlier five parts have been decoded, joined in order, and encoded once as a single Base64 file.

## Use the bundle

1. Download **BASE64_TOOLS_COMPLETE.base64.txt** using the link above or GitHub's **Download raw file** control.
2. Decode the entire file once from Base64 to UTF-8 plain text, then pass that text to your module-import script.
3. Import all 58 modules together into a backup copy of your Access database, replacing existing same-named modules without numbered duplicates.
4. Choose **Debug > Compile** in the VBA editor, then run **SetupLauncher**.

Each complete module uses these exact markers:

```text
'#MODULE ModuleName
(module code)
'#END ModuleName
```

Marker names are actual VBA module names, with no .bas extension or byte count. The old byte-counted v2 Access decoder does not read this format. Both **modLCNReview** (LCNReviewInvalidate) and **modVariantContext** (VC_Setting) are included.

## Included behavior

The tool builds engineering LCN proposals from EBOM/PLMXML, supports vehicle and LCN masters, and retains existing SLIC LCNs by default in protected comparison/output. Changes require field-specific selections. It also supports independent PLMXML LCN corrections and carrying saved LCNs into a newer XML.

The hierarchy update removes the added level-jump stop and uses unique Parent/TCUID evidence when available, with Product Structure Children values available in the hierarchy notices report.

## Validation

The complete source passed **107 assertions in Microsoft Access on Windows 11** on September 15, 2026: 42 SLIC regression checks, 35 variant/master checks and 30 hierarchy checks. The single-file packaging was verified by exact Base64 round trip, complete module coverage and source comparison; no VBA logic changed in this combination.

The reported row-235/1995 hierarchy cases were tested using constructed fixtures. The actual user EBOMs and production SLIC data were unavailable for those tests. Complete cross-ALC ownership, XF/HO usable-on generation and re-home transactions are outside this release's implemented scope.

## File verification

- Encoded file size: **1,225,920 bytes**.
- Decoded bundle size: **919,439 bytes**.
- Complete modules: **58**, each present exactly once.
- SHA-256 of the encoded download: `02ed2fc870bf3a4e8f19e21f3e8ac60edbc5f7a3f52353ac4fd4a56f4183fd5a`.

This is the requested single large GitHub download. The smaller A-E transport parts remain a separate delivery format.
