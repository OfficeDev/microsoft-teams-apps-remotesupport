---
page_type: sample
languages:
- csharp
products:
- office-teams
description: Microsoft Teams bot and messaging extension to search & report incidents and connect with specialists immediately
urlFragment: microsoft-teams-apps-remotesupport
---

# Remote Support App Template

| [Documentation](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/Home) | [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/Deployment-Guide) | [Architecture](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/Solution-Overview) |
| ---- | ---- | ---- |

Most organizations have a team of remote individuals providing support to employees across the organization often distributed geographically. Support and collaboration in such instances is often ad-hoc, sub-optimal and inefficient. Common incumbent solutions include shared email-inbox where employees send in requests; a SharePoint site where requests are submitted; calling a dedicated helpline,  email or chat based messaging systems with a dedicated point person etc.
 
Remote Support bot provides all end users (internal users seeking help from a central team) an easy interface (bot) right within Microsoft Teams to:

- Submit requests for support
- Edit/withdraw requests
- Notify end users about the status of their request
- Escalate to a group chat that connects them immediately with an expert allowing real time video/screen-sharing ability
- Route incoming requests in real-time to a specific/Teams channel which allows the members of the channel an easy interface (a bot within their teams/channel) to:
- See in real-time all incoming requests with associated details
- Start an instant chat or video call with the requester
- Receive and act upon an incoming Teams group chat from the remote requester
- Manage incoming requests within the central team (lightweight ticketing)
- Manage the list of experts who will be on-call to receive incoming Teams group chat requests

![Remote support new request](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/Images/new-request.png)

![Remote support messaging extension](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/Images/support-team-ackcard.jpg)

## Legal notice
This app template is provided under the [MIT License](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/blob/master/LICENSE) terms.  In addition to these terms, by using this app template you agree to the following:

- You, not Microsoft, will license the use of your app to users or organization. 

- This app template is not intended to substitute your own regulatory due diligence or make you or your app compliant with respect to any applicable regulations, including but not limited to privacy, healthcare, employment, or financial regulations.

- You are responsible for complying with all applicable privacy and security regulations including those related to use, collection and handling of any personal data by your app. This includes complying with all internal privacy and security policies of your organization if your app is developed to be sideloaded internally within your organization. Where applicable, you may be responsible for data related incidents or data subject requests for data collected through your app.

- Any trademarks or registered trademarks of Microsoft in the United States and/or other countries and logos included in this repository are the property of Microsoft, and the license for this project does not grant you rights to use any Microsoft names, logos or trademarks outside of this repository. Microsoft’s general trademark guidelines can be found [here](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general.aspx).

- If the app template enables access to any Microsoft Internet-based services (e.g., Office365), use of those services will be subject to the separately-provided terms of use. In such cases, Microsoft may collect telemetry data related to app template usage and operation. Use and handling of telemetry data will be performed in accordance with such terms of use.

- Use of this template does not guarantee acceptance of your app to the Teams app store. To make this app available in the Teams app store, you will have to comply with the [submission and validation process](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.

## Getting started

Begin with the [Solution overview](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/Solution-overview) to read about what the app does and how it works.

When you're ready to try out Remote Support bot, or to use it in your own organization, follow the steps in the [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-remotesupport/wiki/DeployementGuide).

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
