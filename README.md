# spfx-wp-slicklist

## Summary

A SharePoint framework webpart for modern SharePoint (SharePoint Online) meant to present a SharePoint lists data in a simple, but responsonsive table for small data sets that are frequently referenced.
- Ideal for commonly referenced lists with a thousand entries or less
- Loads the entire list into memory and filters client side for very fast lookups
- Allows selection of Lists on other sites on the same tenant
- Allows hiding columns based on three pre-set media queries for mobile, tablet and desktop
- Column filter fields are determined by column data type

[picture of the solution in action, if possible]

## Usage
1. Create a SharePoint list with as many columns as necessary.
2. Internal column names are used for table headers.
3. Column titles (what you rename the columns to after creating them) are used for header tooltips.
3. Unless you want the first column to be named "Title" simply hide that column.
4. By default the list will be displayed as 5 columns wide on mobile, 8 columns wide on tablet, and 10 columns wide on desktop. These can be adjusted for each table however. 
5. There are four field types used for filtering:
    1. Date picker for date type columns
    2. Select menu for boolean type columns, defaulting to nothing selected.
    3. Select menu for choice type columns, defaulting to nothing selected.
    4. Text field for all other column types.
6. Columns added beyond the Desktop Visible Columns attribute will not show in the webpart list but will show in the modal popup. If you do not wish a column to show anywhere in the webpart, simply hide if from the All Items view in the SharePoint list.

## Used SharePoint Framework Version

| :warning: Important          |
|:---------------------------|
| Every SPFx version is only compatible with specific version(s) of Node.js. In order to be able to build this sample, please ensure that the version of Node on your workstation matches one of the versions listed in this section. This sample will not work on a different version of Node.|
|Refer to <https://aka.ms/spfx-matrix> for more information on SPFx compatibility.   |

![version](https://img.shields.io/badge/version-1.18.0-green.svg)
![Node.js v16](https://img.shields.io/badge/Node.js-v16-green.svg) 
![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | February 25, 2024 | Initial Version |

## Minimal Path to Awesome

- Clone or download this repository
- Run in command line:
  - `npm install` to install the npm dependencies
  - `gulp serve` to display in Developer Workbench (recommend using your tenant workbench so you can test with real lists within your site)
- To package and deploy:
  - Use `gulp bundle --ship` & `gulp package-solution --ship`
  - Add the `.sppkg` to your SharePoint App Catalog

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---