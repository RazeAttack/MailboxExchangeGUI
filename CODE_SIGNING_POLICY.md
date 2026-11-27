# Code signing policy

Free code signing provided by [SignPath.io](https://about.signpath.io), certificate by [SignPath Foundation](https://signpath.org).

## Scope

Only the following release artifact is signed:

| File | Description |
|---|---|
| `MailboxAccess.exe` | Launcher with the embedded script `Get-UserMailboxAccess.ps1`, built by `Build-MailboxAccessExe.ps1` |

Signed binaries are built exclusively by the GitHub Actions workflow
[`.github/workflows/release.yml`](.github/workflows/release.yml) on GitHub-hosted runners
from the tagged source code of this repository (`vX.Y.Z`). Nothing is signed from local builds.

The build uses [ps2exe](https://github.com/MScholtes/PS2EXE) (MIT license) to compile the launcher.
All source code and build scripts are maintained in this repository.

## Team roles

| Role | Members |
|---|---|
| Authors (committers) | [@RazeAttack](https://github.com/RazeAttack) |
| Reviewers | [@RazeAttack](https://github.com/RazeAttack) – reviews all external contributions (pull requests) |
| Approvers | [@RazeAttack](https://github.com/RazeAttack) – manually approves every signing request |

All team members use multi-factor authentication for GitHub and SignPath.

## Release process

1. Version and changelog are updated in `src\00-Header.ps1` and `CHANGELOG.md`
2. CI checks pass (syntax, encoding, PSScriptAnalyzer)
3. A tag `vX.Y.Z` is pushed → the release workflow builds `MailboxAccess.exe`
4. The workflow submits a signing request to SignPath; an approver checks the build and approves it manually
5. The signed EXE is published as a GitHub release together with `SHA256SUMS.txt`

## Privacy policy

This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it.

Network connections made by the program:

| Destination | Purpose | Can be disabled |
|---|---|---|
| Microsoft Exchange Online / Microsoft Entra ID | Sign-in and reading/changing mailbox permissions of the user's own tenant – the core function of the tool | – (required) |
| `api.github.com` | Check for a newer release (only the public release information is requested, no data about the user or tenant is sent) | Yes: parameter `-NoUpdateCheck` |

All other data (permission cache, logs, exports) is stored locally in `%LOCALAPPDATA%\EXOMailboxAccess`.

## System changes and uninstallation

The program does not install anything and does not change system settings.

- The EXE writes the script and icon to `%LOCALAPPDATA%\EXOMailboxAccess` on every start.
- Cache, logs and exports are stored in the same folder.

**Uninstall:** delete `MailboxAccess.exe` and the folder `%LOCALAPPDATA%\EXOMailboxAccess`:

```powershell
Remove-Item "$env:LOCALAPPDATA\EXOMailboxAccess" -Recurse -Force
```
