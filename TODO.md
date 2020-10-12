- Add parameter check : if $localPath contains " ' or :, it should fail with an explicit error: bad interpretation of parameter by opcon. Missing escape character \
- Improve winscp output formatting (for now, all output is on one line, which is not clear)
- Additional error checking

- Test & validate S3 protocol
- Test & validate HTTP(S) protocol