# Permissions Manager Mk. I

Created and Developed by : Leo-RD

## Summary
An integrated SharePoint solution designed to simplify permissions management for department heads across an organization’s various divisions (finance, IT, HR, etc.).
Allows users to add and remove permissions of any type.
Comprehensive logging system and email alerts for changes via an automated Power Automate workflow running in the background.

Uses SPFx (SharePoint Framework) Web Parts, allowing the tool to be directly embedded in a SharePoint site while adapting to the context of the currently logged-in user.
Main program written in React TypeScript.
Uses PnPjs.
Uses Node.js/npm.


## Used SharePoint Framework Version
SPFx
![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

**Nodejs v22** 
  - nodejs.org
  - Check installation : `node --version` ; `npm --version`

**SPFx toolchain**
  - in solution folder, run : `npm install @rushstack/heft yo @microsoft/generator-sharepoint --global`

**VSCode or any IDE**

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 0.1     | September 16, 2026  | MVP  |
| 1.0     | ???? ??, 2026 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Quick Start guide

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install -g @rushstack/heft`
  - `npm install`
  - `heft start`

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

## To deploy solution as .sppkg package

- Ensure that you are at the solution folder
- in the commmand-line run:
  - `npx heft test --clean --production`
  - `npx heft package-solution --production`
  - The .sppkg file will be available in `[projectname]\sharepoint\solution`


## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
