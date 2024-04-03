# Overview of Copilot Communication Manifesto Plugin

This demo is based on the search-based [message extension](https://docs.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=nodejs) sample that ships with Teams Tookit.

It returns a fixed set of principles that define good email behavoir. When coupled with Copilot's ability to index sent mail this plugin enables users to be notified when the emails they are sending do not align with their self-defined manifesto for communication.

## Get started with the template

For a full walkthrough of how this sample works, or to start with the original search-based extension and make changes to end up here, there's a walkthrough video at: https://youtu.be/T57cLKC0WYc

## What's included in the template

| Folder        | Contents                                     |
| ------------- | -------------------------------------------- |
| `.vscode/`    | VSCode files for debugging                   |
| `appPackage/` | Templates for the Teams application manifest |
| `env/`        | Environment files                            |
| `infra/`      | Templates for provisioning Azure resources   |
| `src/`        | The source code for the search application   |

The following files can be customized and demonstrate an example implementation to get you started.

| File               | Contents                                                                       |
| ------------------ | ------------------------------------------------------------------------------ |
| `src/searchApp.js` | Handles the business logic for this app to return the communication manifesto. |
| `src/index.js`     | `index.js` is used to setup and configure the Message Extension.               |

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Teams Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                     |

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Teams Toolkit CLI](https://aka.ms/teams-toolkit-cli/debug)
- [Preview the app on mobile clients](https://github.com/OfficeDev/TeamsFx/wiki/Run-and-debug-your-Teams-application-on-iOS-or-Android-client)
