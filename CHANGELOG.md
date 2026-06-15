<h1>Release Notes</h1>

## Recent Changes

### Documentation
- 📝 Refresh roadmap docs and add code of conduct ([#20](https://github.com/mindoffwork/mindoff-dataport/pull/20))
- 📝 Add roadmap and improve project documentation ([#19](https://github.com/mindoffwork/mindoff-dataport/pull/19))

## v1.0.0

### Documentation
- 📝 Add contributing guide and MkDocs documentation setup ([#17](https://github.com/mindoffwork/mindoff-dataport/pull/17))

### Internal
- 🔒 Tighten stable export contract before release ([#18](https://github.com/mindoffwork/mindoff-dataport/pull/18))

## v0.6.1

### Fixes
- 🐛 Fix PDF height-flush dropping multi-row merges at chunk boundaries ([#16](https://github.com/mindoffwork/mindoff-dataport/pull/16))

## v0.6.0

### Enhancements
- 🚀 Enhance dataframe shift handling and validation for merges ([#15](https://github.com/mindoffwork/mindoff-dataport/pull/15))

## v0.5.0

### Enhancements
- ♻ Reorganize example scripts and outputs ([#14](https://github.com/mindoffwork/mindoff-dataport/pull/14))
- 🧪 Increase renderer and contract test coverage with lean branch tests ([#13](https://github.com/mindoffwork/mindoff-dataport/pull/13))
- ⚡ Enhance PDF and XLSX rendering with new fast paths and tests ([#12](https://github.com/mindoffwork/mindoff-dataport/pull/12))

## v0.4.0

### Internal
- 🚚 Organize mindoff_dataport into extract, compile, and render subpackages ([#8](https://github.com/mindoffwork/mindoff-dataport/pull/8))

### Miscellaneous
- ⏪ Revert #8 -- Organize mindoff_dataport into extract, compile, and render subpackages ([#9](https://github.com/mindoffwork/mindoff-dataport/pull/9))

### Enhancements
- ⚡ Speed up XLSX repeat streaming exports ([#10](https://github.com/mindoffwork/mindoff-dataport/pull/10))

### Features
- ✨ Add support for repeating dataframe headers in PDF and tests ([#11](https://github.com/mindoffwork/mindoff-dataport/pull/11))

## v0.3.0

### Fixes
- 🐞 Fix repeat-sheet dataframe-content row offsets in streaming export ([#6](https://github.com/mindoffwork/mindoff-dataport/pull/6))
- 🐞 Fix dataframe shifting issue with repeat blocks ([#4](https://github.com/mindoffwork/mindoff-dataport/pull/4))

### Enhancements
- ✨ Expand style extraction and rendering support ([#5](https://github.com/mindoffwork/mindoff-dataport/pull/5))

### Features
- ✨ Add page break extraction and export support ([#7](https://github.com/mindoffwork/mindoff-dataport/pull/7))

## v0.2.0

### Fixes
- Add configurable dataframe collision shifting during compile
- 🔧 Improve XLSX streaming layout handling ([#1](https://github.com/mindoffwork/mindoff-dataport/pull/1))

### Documentation
- 📝 Refresh README documentation ([#2](https://github.com/mindoffwork/mindoff-dataport/pull/2))

### Features
- ✨ Add configurable collision shifting and PDF chunking coverage ([#3](https://github.com/mindoffwork/mindoff-dataport/pull/3))

## v0.1.0

### Enhancements
- ✨ Add parquet-backed PDF and XLSX export support
- ✨ Add hybrid parquet streaming bundle
- ✨ Add streaming export support
- ✨ Update renderer and streaming
- ✨ Update export API and guidance

### Internal
- ♻️ Refactor export pipeline
- ♻️ Rename package to mindoff_dataport
- ♻️ Move tests into src/tests
- ♻️ Trim examples to core features only
- ♻️ Promote mo_dataport API alias and expand package metadata
- 🔧 Refresh project setup
- 🙈 Move example outputs under ignored folder

### Documentation
- 📝 Refresh README and example import alias
