# Search issue for a specific document

## Summary
This feature helps user diagnose the issue when a specific document does not appear in the search results

Enter the affected site url and select the affected list/library. Then enter the full URL/path of the affected document. Click 'Check Search Document'.

<img src=./asset/NoCrawl.JPG>

```


### Deploy to non-script sites / modern team sites
By default this web part is not allowed on no-script sites, as it allows execution of arbitrary script. This is by design as from a security and governance perspective you might not want arbitrary script added to your pages. This is typically something you want control over.

If you however want to allow the web part for non-script sites like Office 365 Group modern team sites you have to edit `ScriptEditorWebPart.manifest.json` with the following change:
```
"requiresCustomScript": false
```

### Deploy tenant wide
By default you have to install this web part per site collection where you want it availble. If you want it enabled by default on all sites you have to edit `package-solution.json` with the following change:
```
"skipFeatureDeployment": true
```

In order to make it available to absolutely all sites you need apply the _Deploy to non-script sites / modern team site_ change as well.


## Compatibility

![SPFx 1.10](https://img.shields.io/badge/SPFx-1.10.0-green.svg) 
![Node.js LTS 10.x](https://img.shields.io/badge/Node.js-LTS%2010.x%20%7C%20LTS%208.x-green.svg) 
![SharePoint Online](https://img.shields.io/badge/SharePoint-Online-yellow.svg) 
![Teams N/A: Untested with Microsoft Teams](https://img.shields.io/badge/Teams-N%2FA-lightgrey.svg "Untested with Microsoft Teams") 
![Workbench Local | Hosted](https://img.shields.io/badge/Workbench-Local%20%7C%20Hosted-green.svg)

## Applies to

* [SharePoint Framework Release GA](https://blogs.office.com/2017/02/23/sharepoint-framework-reaches-general-availability-build-and-deploy-engaging-web-parts-today/)
* [Office 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)

## Solution

Solution|Author(s)
--------|---------
react-script-editor | Mikael Svenson ([@mikaelsvenson](http://www.twitter.com/mikaelsvenson), [techmikael.com](techmikael.com))

## Version history

Version|Date|Comments
-------|----|--------
1.0|March 10th, 2017|Initial release
1.0.0.1|August 8th, 2017|Updated SPFx version and CSS loading
1.0.0.2|October 4th, 2017|Updated SPFx version, bundle Office UI Fabric and CSS in webpart
1.0.0.3|January 10th, 2018|Updated SPFx version, added remove padding property and refactoring
1.0.0.4|February 14th, 2018|Added title property for edit mode and documentation for enabling the web part on Group sites / tenant wide
1.0.0.5|March 8th, 2018|Added support for loading scripts which are AMD/UMD modules. Added support for classic _spPageContextInfo. Refactoring.
1.0.0.6|March 26th, 2018|Fixed so that AMD modules don't detect `define`, and load as non-modules.
1.0.0.7|May 23rd, 2018|Added supportsFullBleed to manifest.
1.0.0.8|May 23rd, 2018|Updated SPFx to v1.5.1, made editor load dynamically to reduce runtime bundle size, fixed white space issue.
1.0.0.9|Aug 22, 2018|Improved bundle size
1.0.0.10|Jan 16th, 2019|Fix for removing of web part padding
1.0.0.11|March 18th, 2019|Fix for re-loading of script on smart navigation
1.0.0.12|April 15th, 2019|Re-fix for pad removal of web part
1.0.0.13|July 1th, 2019|Downgrade to SPFx v1.4.1 to support SP2019
1.0.0.14|Oct 13th, 2019|Added resolve to fix pnpm issue. Updated author info.
1.0.0.15|Mar 16th, 2020|Upgrade to SPFx v1.10.0. Add support for Teams tab. Renamed package file.
1.0.0.16|April 1st, 2020|Improved how script tags are handled and cleaned up on smart page navigation.
1.0.17.0|January 29th, 2021|Changed versioning to 3 parts. Updated npm packages, restructured documentation, minor change to webpack analyzer setup.

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome
### Local testing

- Clone this repository
- In the command line run:
  - `npm install`
  - `gulp serve`

### Deploy
* gulp clean
* gulp bundle --ship
* gulp package-solution --ship
* Upload .sppkg file from sharepoint\solution to your tenant App Catalog
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog
* Add the web part to a site collection, and test it on a page
## Features
This web part illustrates the following concepts on top of the SharePoint Framework:

- Re-use existing JavaScript solutions on a modern page
- Office UI Fabric
- React

<img src="https://telemetry.sharepointpnp.com/sp-dev-fx-webparts/samples/react-script-editor" />