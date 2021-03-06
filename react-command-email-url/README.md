## react-command-email-url

In SharePoint there used to be a 'copy shortcut' option in the right-click menu on a document. This feature is not available currently. Now we need to go to the share sub-menu to get a link to the document, but what is offered there is the docidredir link, not the full path.

This SPFx extension opens a dialog where the user can see the document full path. The Email Link button helps the user to share the link via default email client.

![react-command-email-url](./assets/sample.png)
![react-command-email-url](./assets/email.png)

## Used SharePoint Framework Version 
SPFx v1.6

## Applies to

* [SharePoint Framework](http://dev.office.com/sharepoint/docs/spfx/sharepoint-framework-overview)
* [Office 365 developer tenant](http://dev.office.com/sharepoint/docs/spfx/set-up-your-developer-tenant)

Solution|Author(s)
--------|---------
react-command-email-url|Joseph Velliah (SPRIDER, @sprider)

## Version history

Version|Date|Comments
-------|----|--------
1.0|October 20, 2018|Initial release

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Minimal Path to Awesome

- Clone [this](https://github.com/sprider/spfx-extensions/tree/master/react-command-email-url) repository
  - Open serve.json file under config folder. Update PageUrl to the URL of the list you wish to test.
![react-command-email-url](./assets/servejson.png)
- Move to folder where this readme exists
- In the command window run:
  - `npm install`
  - `gulp serve`

## Features
This project contains SharePoint Framework extensions that illustrates the following features:
* Command extension
* Office UI Fabric React

> Notice. This sample is designed to be used in debug mode and does not contain automatic packaging setup for the "production" deployment.


