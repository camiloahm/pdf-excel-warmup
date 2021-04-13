# pdf-excel-warmup

This POC converts multiple PDFs into excel sheets. The structure pf the PDFs is static, the code doesn't look that nice
but is useful to understand how to transform pdfs into xlsx.

## Packaging executable

It uses appassembler, so you can generate an executable script with

```
mvn package appassembler:assemble
```

then you just share the appassembler folder inside the target