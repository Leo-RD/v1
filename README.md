# permissions-manager-mk-i

## Summary
Solution intégrée dans SharePoint ayant pour objectif de simplifier la gestion des droits, à destination des directeurs des différents poles d'une entité (financier, informatique, RH etc)
Permet d'ajouter, de supprimer des droits de tout type
Système de logs complet, et alerte par e-mail des modifications à l'aide d'un flux automatisé Power Automate qui fonctionne en arrière-plan.

Utilisation de WebPart SPFx (SharePoint Framework), permettant d'inclure directement l'outil dans un site sharepoint en s'adaptant au contexte de l'utilisateur déjà connecté.
Programme principal écrit en TypeScript.
Utilisation de PnPjs.
Utilisation de NodeJS/npm.
Utilisation de React.

## Used SharePoint Framework Version
SPFx 1.22+
![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Nodejs v22
> SPFx toolchain
> VSCode, or any IDE

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install -g @rushstack/heft`
  - `npm install`
  - `heft start`

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

> To deploy solution as .sppkg package

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
