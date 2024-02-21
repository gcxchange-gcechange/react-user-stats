# User Stats

## Summary
The web component provides comprehensive statistics including:

- Total user count and active users.
- Detailed monthly user reports.
- Total number of communities.
- Distribution of communities by department.
- Breakdown of community membership count.
- Distribution of members across communities.
- Breakdown of community site storage capacity.
- File count breakdown per community site.
- Downloadable CSV file data and source file reports from Azure Active Directory


## Prerequisites
This web part connects to this [this function app](https://github.com/gcxchange-gcechange/appsvc-fnc-dev-userstats).


## Version 
![SPFX](https://img.shields.io/badge/SPFX-1.17.4-green.svg)
![Node.js](https://img.shields.io/badge/Node.js-v16.3+-green.svg)

## Applies to
- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Version history

Version|Date|Comments
-------|----|--------
1.0|December 09, 2021 |Initial release
1.1|February 20, 2024 |Upgrade to 1.17.4

## Minimal Path to Awesome
- Clone this repository
- Ensure that you are at the solution folder
- To install the dependencies, in the command-line run:
  - **npm install**
- To debug in the front end:
  - go to the `serve.json` file and update `initialPage` to `https://your-domain-name.sharepoint.com/_layouts/15/workbench.aspx`
  - In the command-line run:
    - **gulp serve**
    - You will need to add your client id and azure function to the clientId and url classs members in the SCW.tsx file.
- To deploy:
  - In the command-line run:
    - **gulp clean**
    - **gulp bundle --ship**
    - **gulp package-solution --ship**
  - Add the webpart to your tenant app store
- Approve the web API permissions
- Add the Webpart to a page
- Modify the property pane according to your requirements
 

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
