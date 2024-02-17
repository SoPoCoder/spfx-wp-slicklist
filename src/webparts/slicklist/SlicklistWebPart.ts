import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from '../../utils/pnpjs-config';
import Slicklist, { ISlicklistProps } from './components/Slicklist';
import "@pnp/sp/search";
import { ISearchQuery, SearchResults } from "@pnp/sp/search";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'SlicklistWebPartStrings';

export interface ISlicklistWebPartProps {
  table1Title: string;
  table1SiteURL: string;
  table1ListName: string;
  visibleColsMobile: number;
  visibleColsTablet: number;
  visibleColsDesktop: number;
  displayMode: DisplayMode;
}

export interface ISPSite {
  SPSiteUrl: string;
  Title: string;
}
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
}

/*-----------------------------------------------------
Todo List:
1. allow selecting List View, default to All Items
3. allow adding second lookup table for additional information in modal popup
4. add property for toggling between show all/none by default (will require paging)
5. test for desktop, tablet, mobile
-----------------------------------------------------*/
export default class SlicklistWebPart extends BaseClientSideWebPart<ISlicklistWebPartProps> {

  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = getSP(this.context);
    this.properties.visibleColsMobile = 5;
    this.properties.visibleColsTablet = 8;
    this.properties.visibleColsDesktop = 10;
  }

  public render(): void {
    const element: React.ReactElement<ISlicklistProps> = React.createElement( Slicklist,
      {
        table1Title: this.properties.table1Title,
        table1SiteURL: this.properties.table1SiteURL,
        table1ListName: this.properties.table1ListName,
        visibleColsMobile: this.properties.visibleColsMobile,
        visibleColsTablet: this.properties.visibleColsTablet,
        visibleColsDesktop: this.properties.visibleColsDesktop,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.table1Title = value;
        },
        onConfigure: () => {
          this.context.propertyPane.open();
        },
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /* ------------------------------------------------------------ */
  /* --------------------- Properties Panel --------------------- */
  /* ------------------------------------------------------------ */

  private _siteSelectOptions: IPropertyPaneDropdownOption[] = [];
  private _listSelectOptions: IPropertyPaneDropdownOption[] = [];

  private _siteNameDropdownDisabled = true;
  private _listNameDropdownDisabled = true;

  private async _getSiteNames(): Promise<boolean> {
    //const queryPath: string = "path:" + window.location.hostname + "/sites/ ";
    const results: SearchResults = await this._sp.search(<ISearchQuery>{
      Querytext: "contentclass:STS_Site",
      RowLimit: 500,
      SelectProperties: ["SPSiteUrl", "Title"]
    });
    const sites = results.PrimarySearchResults;
    this._siteSelectOptions = sites.map((item: ISPSite) => {
      return {
        key: item.SPSiteUrl,
        text: item.Title
      };
    }).sort((a, b) => a.text.localeCompare(b.text));
    this._siteNameDropdownDisabled = false;
    return true;
  }

  private async _getlistNames(): Promise<boolean> {
    if (this.properties.table1SiteURL) {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.properties.table1SiteURL}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1);
      const lists: Promise<ISPLists> = await response.json();
      this._listSelectOptions = (await lists).value.map((item: ISPList) => {
        return {
          key: item.Title,
          text: item.Title
        };
      });
      this._listNameDropdownDisabled = false;
      return true;
    }
    return false;
  }

    // fired when the properties panel is opened
    protected onPropertyPaneConfigurationStart(): void {
      if (this.properties.table1SiteURL.trim().length < 1) {
        this.properties.table1SiteURL = this.context.pageContext.web.absoluteUrl;
      }
      this._getSiteNames().then((result) => {
        if (result){
          this._getlistNames().then((result) => {
            if (result){
              this.context.propertyPane.refresh();
              this.render();
            }
          }).catch((error: Error) => {throw error});
        }
      }).catch((dsferror: Error) => {/* console.log(error) */})
    }
  
    // fired when the properties panel is changed (see https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/use-cascading-dropdowns-in-web-part-properties)
    protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | undefined, newValue: string | undefined): Promise<void> {
  
      if (propertyPath === "table1SiteURL") {
        await this._getlistNames();
        this.context.propertyPane.refresh();
        this.render();
      } else {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        this.context.propertyPane.refresh();
        this.render();
      }
    }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.Table1GroupName,
              groupFields: [

                PropertyPaneTextField('table1Title', {
                  label: strings.Table1TitleFieldLabel,
                }),
                PropertyPaneDropdown('table1SiteURL', {
                  label: strings.Table1SiteNameFieldLabel,
                  options: this._siteSelectOptions,
                  selectedKey: this.properties.table1SiteURL,
                  disabled: this._siteNameDropdownDisabled,
                }),
                PropertyPaneDropdown('table1ListName', {
                  label: strings.Table1ListNameFieldLabel,
                  options: this._listSelectOptions,
                  selectedKey: this.properties.table1ListName,
                  disabled: this._listNameDropdownDisabled,
                }),
                PropertyPaneSlider('visibleColsMobile', {
                  label: strings.Table1VisibleColsMobileFieldLabel,
                  min: 1,
                  max: 20,
                  value: this.properties.visibleColsMobile
                }),
                PropertyPaneSlider('visibleColsTablet', {
                  label: strings.Table1VisibleColsTabletFieldLabel,
                  min: 1,
                  max: 20,
                  value: this.properties.visibleColsTablet
                }),
                PropertyPaneSlider('visibleColsDesktop', {
                  label: strings.Table1VisibleColsDesktopFieldLabel,
                  min: 1,
                  max: 20,
                  value: this.properties.visibleColsDesktop
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
