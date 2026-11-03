# Changelog

All notable changes to this project are documented here. This file is maintained
automatically by [semantic-release](https://github.com/semantic-release/semantic-release)
on every release to `main`.

## [0.4.0](https://github.com/bauer-group/IP-PDFFolderPrint/compare/v0.3.0...v0.4.0) (2026-11-03)

### 🚀 Features

* **gui:** redesigned German-first UI with form grid and clean visual hierarchy ([ea9700a](https://github.com/bauer-group/IP-PDFFolderPrint/commit/ea9700a74f0949bcce6cbeacf52ed520191f9144))

## [0.3.0](https://github.com/bauer-group/IP-PDFFolderPrint/compare/v0.2.1...v0.3.0) (2026-11-03)

### 🚀 Features

* **app:** add --yes flag, fail on unreadable PDFs, GUI delay, and branded EXE icon ([274a63b](https://github.com/bauer-group/IP-PDFFolderPrint/commit/274a63bcaf2e769089c593326355e7b572f55a37)), references [#FF8500](https://github.com/bauer-group/IP-PDFFolderPrint/issues/FF8500)
* **app:** cancel, print report, printer status, retries, natural sort, settings persistence ([27183d4](https://github.com/bauer-group/IP-PDFFolderPrint/commit/27183d4840ad36e8548e00d3f153a72ab6203cd8))

### 🐛 Bug Fixes

* **ci:** added the missing permissions block ([6e8af5b](https://github.com/bauer-group/IP-PDFFolderPrint/commit/6e8af5bab8ae9485e99ef69a0a7bb681655922eb))

### ♻️ Code Refactoring

* restructure into src package with tests, pinned deps and CI pipeline ([a24b2d1](https://github.com/bauer-group/IP-PDFFolderPrint/commit/a24b2d1fbdcafd856c58ffeefcda098d22280609))

### 🔧 Maintenance

* **ci:** removed issue AI summary workflow ([b920e0d](https://github.com/bauer-group/IP-PDFFolderPrint/commit/b920e0d2743d839a85dbebb1ce286c4421268483)), references [bauer-group/automation-templates#105](https://github.com/bauer-group/automation-templates/issues/105)
* **ci:** removed redundant teams notification ([d570fb6](https://github.com/bauer-group/IP-PDFFolderPrint/commit/d570fb6dcf077bfc6e9923edff924c454886840a))
* **ci:** use Python 3.14 for CI lint, tests and EXE build ([1b1e32f](https://github.com/bauer-group/IP-PDFFolderPrint/commit/1b1e32f1c48caac8f15254cd1c86e6d9329b3f58))
* **codeowners:** reassigned ownership to core team [skip ci] ([7bcc1ab](https://github.com/bauer-group/IP-PDFFolderPrint/commit/7bcc1ab4bf5ddb0c34888a3448a1896ec356c18b))

## [0.2.1](https://github.com/bauer-group/IP-PDFFolderPrint/compare/v0.2.0...v0.2.1) (2026-04-02)

### 🐛 Bug Fixes

* **build:** forced UTF-8 output encoding on Windows ([79f7fa4](https://github.com/bauer-group/IP-PDFFolderPrint/commit/79f7fa4b81145da61b8f8e0eea3fcb2f18f8cfe2))

## [0.2.0](https://github.com/bauer-group/IP-PDFFolderPrint/compare/v0.1.1...v0.2.0) (2026-04-02)

### 🚀 Features

* **build:** added --ci flag to skip venv and use system Python ([33adff8](https://github.com/bauer-group/IP-PDFFolderPrint/commit/33adff8f61d7974b0589a53ac980788a758169a4))

## [0.1.1](https://github.com/bauer-group/IP-PDFFolderPrint/compare/v0.1.0...v0.1.1) (2026-04-02)

### 🐛 Bug Fixes

* **ci:** added requirements.txt for pip cache resolution ([19f3605](https://github.com/bauer-group/IP-PDFFolderPrint/commit/19f3605acf7313201ec3ace439098d71de3d050c))
* **ci:** removed pip cache from EXE build workflow ([bea2b22](https://github.com/bauer-group/IP-PDFFolderPrint/commit/bea2b224171900d584d068b7b6ce4fc59795c6f0))
* **ci:** switched EXE build trigger to workflow_run ([e543faf](https://github.com/bauer-group/IP-PDFFolderPrint/commit/e543faf56dfac8b56739abf0c444a4be5e587d0f))

## [0.1.0](https://github.com/bauer-group/IP-PDFFolderPrint/compare/v0.0.0...v0.1.0) (2026-04-02)

### 🚀 Features

* **app:** added PDF batch printing tool with GUI and CLI ([6887c97](https://github.com/bauer-group/IP-PDFFolderPrint/commit/6887c97883df6c7f5f005be28aec49c11c9e4c2d))
