# My Apps launcher for Microsoft Teams
 
A tiny Microsoft Teams personal app that works as a doorway. Select **My Apps** on the Teams left rail, then **Open My Apps**, and [Microsoft My Apps](https://myapplications.microsoft.com/) opens in your web browser.
 
> [!IMPORTANT]
> **Community sample. Use at your own risk.** This project is provided "as is", without warranty of any kind, under the [MIT License](LICENSE). It is not an official Microsoft product, and it is not supported by Microsoft or under any Microsoft support program or service. Test it in a non-production environment and review it with your own security, privacy and legal teams before you use it. Read the full [legal notice](#legal-notice) below.
 
## Get the kit
 
| I want to... | Do this |
|---|---|
| Download the files | Get the ZIP from the [latest release](https://github.com/ShanePoint-com/JavaScript.MS-Teams.MyApps-Launcher/releases/latest), or select **Code** > **Download ZIP** on this page. |
| Start my own copy | Select **Use this template** > **Create a new repository**. You get a clean copy in your own account, and it can be private. |
| Suggest a fix | Fork this repository, make your change, and open a pull request. |
| Work with Git | `git clone https://github.com/ShanePoint-com/JavaScript.MS-Teams.MyApps-Launcher.git` |
 
Direct download of the latest packaged kit: https://github.com/ShanePoint-com/JavaScript.MS-Teams.MyApps-Launcher/releases/latest/download/MyApps_Teams_Launcher_Starter_Kit.zip
 
## What it does
 
```text
Teams left rail  ->  My Apps tab  ->  Open My Apps  ->  My Apps opens in your web browser
```
 
- Uses the Teams JavaScript library (TeamsJS) `app.openLink()` to open My Apps outside Teams, with a normal web link as a backup.
- No bot, no API, no database, no back end, no sign-in code, and no Microsoft Entra app registration.
- The page doesn't collect, store or send personal data.
 
## What's inside
 
| Path | What it is |
|---|---|
| `web/` | The launcher page, styles and logic, Azure Static Web Apps settings, and sample privacy, terms and support pages |
| `appPackage/` | Teams app manifest template (schema 1.30) and the two app icons |
| `scripts/` | Preview the page locally, check the published site, and build the Teams app package |
| `tests/` | Optional automated browser test |
| `QUICKSTART.md` | Short setup steps |
| `LICENSE` | MIT License |
| `SECURITY.md` | How to report a security issue privately |
 
## Requirements
 
- A Microsoft 365 tenant with Microsoft Teams, and permission to upload custom apps (or a Teams admin who can)
- An Azure subscription for Azure Static Web Apps (the Free plan works for testing)
- Node.js LTS to run the helper scripts, and Git if you want to clone
 
## Quick start
 
1. Get the kit (see above) and open the folder in Visual Studio Code.
2. Make your own Teams app ID: run `[guid]::NewGuid()` in PowerShell, then paste the result into `"id"` in `appPackage/manifest.json`.
3. Publish the `web` folder to an Azure Static Web App, then check it: `node scripts/check-website.mjs https://YOUR-HOSTNAME/`
4. Build the Teams package: `node scripts/build-teams-package.mjs --host YOUR-HOSTNAME --org "Your organization"`
5. In Teams, select **Apps** > **Manage your apps** > **Upload an app** and choose `dist/MyAppsTeamsApp-1.0.0.zip`. After testing, a Teams admin can pin it for people with an app setup policy.
 
More detail is in [QUICKSTART.md](QUICKSTART.md).
 
## Security and privacy
 
- HTTPS only, with a strict Content Security Policy that lets only Microsoft Teams, Outlook and Microsoft 365 show the page in a frame.
- TeamsJS loads from Microsoft's content delivery network with a Subresource Integrity (tamper) check.
- No cookies, no analytics and no tracking.
- Please report security issues privately. See [SECURITY.md](SECURITY.md).
 
## Support
 
This is a community sample. You're welcome to open an issue for questions or ideas, but there's no guaranteed response time and no service-level agreement.
 
## Legal notice
 
- **License.** This project is licensed under the [MIT License](LICENSE). As the license says, the software is provided "as is", without warranty of any kind, and the authors and copyright holders aren't liable for any claim, damages or other liability that comes from its use.
- **Not a Microsoft product.** This is an independent community sample. It isn't an official Microsoft product, and Microsoft doesn't support, endorse or warrant it.
- **Views.** This project is shared by its author. It doesn't represent the views or commitments of any employer or customer.
- **Your responsibility.** You're responsible for how you deploy and use this project, including testing, security review, accessibility, privacy, and following your organization's policies and any laws that apply to you. Nothing here is legal, security or compliance advice.
- **Data.** The launcher doesn't collect, store or send personal data. The hosting service you choose (for example, Azure Static Web Apps) may keep standard technical logs under its own terms.
- **Third-party components.** The page loads the Microsoft Teams JavaScript client library (TeamsJS) from Microsoft's content delivery network. TeamsJS is licensed separately by Microsoft under the MIT License. Your use of Microsoft 365, Microsoft Teams, Azure and My Apps is covered by your own agreements with Microsoft.
- **Sample content.** The privacy, terms and support pages in `web/` are samples. Replace them with your organization's approved text before production use.
- **Trademarks.** Microsoft, Microsoft 365, Microsoft Teams, Microsoft Entra and Azure are trademarks of the Microsoft group of companies. Other names may be trademarks of their respective owners. They're used here only to identify the products this sample works with, and their use doesn't imply endorsement. This project doesn't give you any right to use anyone's trademarks or logos.
 
## License
 
Licensed under the [MIT License](LICENSE).
