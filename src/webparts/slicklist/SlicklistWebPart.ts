import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    IPropertyPaneDropdownOption,
    IPropertyPaneGroup,
    PropertyPaneCheckbox,
    PropertyPaneDropdown,
    PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from '../../utils/pnpjs-config';
import SlickList from './components/SlickList';
import "@pnp/sp/search";
import { ISearchQuery, SearchResults } from "@pnp/sp/search";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'SlicklistWebPartStrings';
import { Web } from '@pnp/sp/webs';
import { IFieldInfo } from '@pnp/sp/fields';
import { ISPList, ISPLists, ISPSite, ISlickListProps, ISlicklistWebPartProps } from '../..';

export default class SlicklistWebPart extends BaseClientSideWebPart<ISlicklistWebPartProps> {

    private _sp: SPFI;
    private _visColMin: number = 1;
    private _visColMax: number = 20;

    protected async onInit(): Promise<void> {
        await super.onInit();
        this._sp = getSP(this.context);
    }

    public render(): void {
        const top = this.domElement as HTMLDivElement;
        const slickList: React.ReactElement<ISlickListProps> = React.createElement(SlickList,
            {
                table1Title: this.properties.table1Title,
                table1SiteURL: this.properties.table1SiteURL,
                table1ListName: this.properties.table1ListName,
                table1VisColsMobile: this.properties.table1VisColsMobile,
                table1VisColsTablet: this.properties.table1VisColsTablet,
                table1VisColsDesktop: this.properties.table1VisColsDesktop,
                lookupColumn: this.properties.lookupColumn,
                table2Title: this.properties.table2Title,
                table2SiteURL: this.properties.table2SiteURL,
                table2ListName: this.properties.table2ListName,
                showTable2: this.properties.showTable2,
                table2VisColsMobile: this.properties.table2VisColsMobile,
                table2VisColsTablet: this.properties.table2VisColsTablet,
                table2VisColsDesktop: this.properties.table2VisColsDesktop,
                orderByColumn1: this.properties.orderByColumn1,
                orderByColumn2: this.properties.orderByColumn2,
                orderByColumn3: this.properties.orderByColumn3,
                orderByColumn4: this.properties.orderByColumn4,
                onConfigure: () => { this.context.propertyPane.open(); },
                onTopClick: () => { top.scrollIntoView({ behavior: 'smooth' }); },
                onLookupClick: (value: string) => {
                    const lookupRow = this.domElement.querySelector(`#${CSS.escape(value)}`) as HTMLTableRowElement;
                    lookupRow.scrollIntoView({ behavior: 'smooth' });
                }
            }
        );
        ReactDom.render(slickList, this.domElement);
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
    private _list1SelectOptions: IPropertyPaneDropdownOption[] = [];
    private _list1ColSelectOptions: IPropertyPaneDropdownOption[] = [];
    private _list2SelectOptions: IPropertyPaneDropdownOption[] = [];
    private _list2ColSelectOptions: IPropertyPaneDropdownOption[] = [];

    private _siteNameDropdownDisabled = true;
    private _list1NameDropdownDisabled = true;
    private _lookupColumnDropdownDisabled = true;
    private _list2NameDropdownDisabled = true;
    private _orderByColumn1DropdownDisabled = true;
    private _orderByColumn2DropdownDisabled = true;
    private _orderByColumn3DropdownDisabled = true;
    private _orderByColumn4DropdownDisabled = true;

    private async getSiteNames(): Promise<boolean> {
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

    private async getlistNames(siteURL: string): Promise<IPropertyPaneDropdownOption[]> {
        if (siteURL) {
            const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${siteURL}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1);
            const lists: Promise<ISPLists> = await response.json();
            return (await lists).value.map((item: ISPList) => {
                return {
                    key: item.Title,
                    text: item.Title
                };
            });
        }
        return [];
    }

    private async getColumnChoices(siteURL: string, listName: string): Promise<IPropertyPaneDropdownOption[]> {
        const selectOptions: IPropertyPaneDropdownOption[] = [];
        const web = Web([this._sp.web, siteURL]);
        const columns = await web.lists.getByTitle(listName).fields.filter("ReadOnlyField eq false and Hidden eq false")();
        if (columns.length > 0) {
            columns.map((column: IFieldInfo) => {
                if (column.Title.trim().length > 0) {
                    selectOptions.push({ key: column.InternalName, text: column.Title });
                }
            });
            selectOptions.unshift({ key: "", text: "" }); // add unselected option to the top of the list
            return selectOptions;
        }
        return selectOptions;
    }

    // the following set of functions chain together to set or reset property panel fields
    private setResetTable1ListName(reset:boolean = false): void {
        this.getlistNames(this.properties.table1SiteURL).then((result: IPropertyPaneDropdownOption[]) => {
            if (reset) {
                this._list1SelectOptions = [];
                this._list1NameDropdownDisabled = true;
            }
            if (result) {
                this._list1SelectOptions = result;
                this._list1NameDropdownDisabled = false;
                this.context.propertyPane.refresh();
                if (this.properties.table1SiteURL && this.properties.table1ListName)
                    this.setResetLookupColumn(reset);
            }
        }).catch((error: Error) => { console.log(error) });
    }
    private setResetLookupColumn(reset:boolean = false): void {
        this.getColumnChoices(this.properties.table1SiteURL, this.properties.table1ListName).then((result: IPropertyPaneDropdownOption[]) => {
            if (reset) {
                this.properties.lookupColumn = "";
                this._list1ColSelectOptions = [];
                this._lookupColumnDropdownDisabled = true;
            }
            if (result) {
                this._list1ColSelectOptions = result;
                this._lookupColumnDropdownDisabled = false;
                this.context.propertyPane.refresh();
                this.setResetTable2SiteURL(reset);
            }
        }).catch((error: Error) => { console.log(error) });
    }
    private setResetTable2SiteURL(reset:boolean = false): void {
        if (reset)
            this.properties.table2SiteURL = this.context.pageContext.web.absoluteUrl;
        this.setResetTable2ListName(reset);
    }
    private setResetTable2ListName(reset:boolean = false): void {
        this.getlistNames(this.properties.table2SiteURL).then((result: IPropertyPaneDropdownOption[]) => {
            if (reset) {
                this.properties.table2ListName = "";
                this._list2SelectOptions = [];
                this._list2NameDropdownDisabled = true;
            }
            if (result) {
                this._list2SelectOptions = result;
                this._list2NameDropdownDisabled = false;
                this.context.propertyPane.refresh();
                this.setResetShowTable2(reset);
            }
        }).catch((error: Error) => { console.log(error) });
    }
    private setResetShowTable2(reset:boolean = false): void {
        if (reset) {
            this.properties.showTable2 = false;
        }
        this.setResetTable2Props();
    }
    private setResetTable2Props(): void {
        this.getColumnChoices(this.properties.table2SiteURL, this.properties.table2ListName).then((result: IPropertyPaneDropdownOption[]) => {
            if (result) {
                this._list2ColSelectOptions = result;
                this._orderByColumn1DropdownDisabled = false;
                this._orderByColumn2DropdownDisabled = false;
                this._orderByColumn3DropdownDisabled = false;
                this._orderByColumn4DropdownDisabled = false;
                this.context.propertyPane.refresh();
            }
        }).catch((error: Error) => { console.log(error) });
    }

    // fired when the properties panel is opened
    protected onPropertyPaneConfigurationStart(): void {
        if (this.properties.table1SiteURL.trim().length < 1) {
            this.properties.table1SiteURL = this.context.pageContext.web.absoluteUrl;
        }
        if (this.properties.table2SiteURL.trim().length < 1) {
            this.properties.table2SiteURL = this.context.pageContext.web.absoluteUrl;
        }
        this.properties.table1VisColsMobile = this.properties.table1VisColsMobile || 2;
        this.properties.table1VisColsTablet = this.properties.table1VisColsTablet || 5;
        this.properties.table1VisColsDesktop = this.properties.table1VisColsDesktop || 7;
        this.properties.table2VisColsMobile = this.properties.table2VisColsMobile || 2;
        this.properties.table2VisColsTablet = this.properties.table2VisColsTablet || 5;
        this.properties.table2VisColsDesktop = this.properties.table2VisColsDesktop || 7;
        this.context.propertyPane.refresh();

        this.getSiteNames().then((gotSites) => {
            if (gotSites) {
                this.setResetTable1ListName();
                this.context.propertyPane.refresh();
            }
        }).catch((error: Error) => { console.log(error) });

    }

    // fired when the properties panel is changed
    protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | undefined, newValue: string | undefined): Promise<void> {
        
        if (propertyPath === "table1SiteURL") {
            this.setResetTable1ListName(true);
        } else if (propertyPath === "table1ListName") {
            if (newValue && newValue !== oldValue)
                this.properties.table1Title = newValue;
            this.setResetLookupColumn(true);
        } else if (propertyPath === "lookupColumn") {
            this.setResetTable2SiteURL(true);
        } else if (propertyPath === "table2SiteURL") {
            this.setResetTable2ListName(true);
        } else if (propertyPath === "table2ListName") {
            if (newValue && newValue !== oldValue)
                this.properties.table2Title = newValue;
            this.setResetShowTable2(true);
        } else if (propertyPath === "showTable2") {
            this.setResetTable2Props();
        } else {
            super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        }
        this.context.propertyPane.refresh();
        this.render();
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

        const propertyPaneGroups: IPropertyPaneGroup[] = [
            {
                groupName: strings.Table1GroupName,
                groupFields: [
                    PropertyPaneDropdown('table1SiteURL', {
                        label: strings.Table1SiteNameFieldLabel,
                        options: this._siteSelectOptions,
                        selectedKey: this.properties.table1SiteURL,
                        disabled: this._siteNameDropdownDisabled,
                        
                    }),
                    PropertyPaneDropdown('table1ListName', {
                        label: strings.Table1ListNameFieldLabel,
                        options: this._list1SelectOptions,
                        selectedKey: this.properties.table1ListName,
                        disabled: this._list1NameDropdownDisabled,
                    }),
                    PropertyPaneSlider('table1VisColsMobile', {
                        label: strings.Table1VisibleColsMobileFieldLabel,
                        min: this._visColMin,
                        max: this._visColMax,
                        value: this.properties.table1VisColsMobile
                    }),
                    PropertyPaneSlider('table1VisColsTablet', {
                        label: strings.Table1VisibleColsTabletFieldLabel,
                        min: this._visColMin,
                        max: this._visColMax,
                        value: this.properties.table1VisColsTablet
                    }),
                    PropertyPaneSlider('table1VisColsDesktop', {
                        label: strings.Table1VisibleColsDesktopFieldLabel,
                        min: this._visColMin,
                        max: this._visColMax,
                        value: this.properties.table1VisColsDesktop
                    }),
                    PropertyPaneDropdown('lookupColumn', {
                        label: strings.LookupColumnFieldLabel,
                        options: this._list1ColSelectOptions,
                        selectedKey: this.properties.lookupColumn,
                        disabled: this._lookupColumnDropdownDisabled
                    })
                ]
            }
        ]

        if (this.properties.lookupColumn) {
            propertyPaneGroups.push(
                {
                    groupName: strings.Table2GroupName,
                    groupFields: [
                        PropertyPaneDropdown('table2SiteURL', {
                            label: strings.Table2SiteNameFieldLabel,
                            options: this._siteSelectOptions,
                            selectedKey: this.properties.table2SiteURL,
                            disabled: this._siteNameDropdownDisabled,
                        }),
                        PropertyPaneDropdown('table2ListName', {
                            label: strings.Table2ListNameFieldLabel,
                            options: this._list2SelectOptions,
                            selectedKey: this.properties.table2ListName,
                            disabled: this._list2NameDropdownDisabled,
                        }),
                        PropertyPaneCheckbox('showTable2', {
                            text: strings.ShowTable2FieldLabel,
                            checked: this.properties.showTable2
                        })
                    ]
                }
            )
        }
        
        if (this.properties.lookupColumn && this.properties.showTable2) {
            propertyPaneGroups.push(
                {
                    groupName: strings.Table2AddPropsGroupName,
                    groupFields: [
                        PropertyPaneSlider('table2VisColsMobile', {
                            label: strings.Table2VisibleColsMobileFieldLabel,
                            min: this._visColMin,
                            max: this._visColMax,
                            value: this.properties.table2VisColsMobile
                        }),
                        PropertyPaneSlider('table2VisColsTablet', {
                            label: strings.Table2VisibleColsTabletFieldLabel,
                            min: this._visColMin,
                            max: this._visColMax,
                            value: this.properties.table2VisColsTablet
                        }),
                        PropertyPaneSlider('table2VisColsDesktop', {
                            label: strings.Table2VisibleColsDesktopFieldLabel,
                            min: this._visColMin,
                            max: this._visColMax,
                            value: this.properties.table2VisColsDesktop
                        }),
                        PropertyPaneDropdown('orderByColumn1', {
                            label: strings.Table2OrderByColumn1FieldLabel,
                            options: this._list2ColSelectOptions,
                            selectedKey: this.properties.orderByColumn1,
                            disabled: this._orderByColumn1DropdownDisabled
                        }),
                        PropertyPaneDropdown('orderByColumn2', {
                            label: strings.Table2OrderByColumn2FieldLabel,
                            options: this._list2ColSelectOptions,
                            selectedKey: this.properties.orderByColumn2,
                            disabled: this._orderByColumn2DropdownDisabled
                        }),
                        PropertyPaneDropdown('orderByColumn3', {
                            label: strings.Table2OrderByColumn3FieldLabel,
                            options: this._list2ColSelectOptions,
                            selectedKey: this.properties.orderByColumn3,
                            disabled: this._orderByColumn3DropdownDisabled
                        }),
                        PropertyPaneDropdown('orderByColumn4', {
                            label: strings.Table2OrderByColumn4FieldLabel,
                            options: this._list2ColSelectOptions,
                            selectedKey: this.properties.orderByColumn4,
                            disabled: this._orderByColumn4DropdownDisabled
                        })
                    ]
                }
            )
        }

        return {
            pages: [
                {
                    header: { description: strings.PropertyPaneDescription },
                    groups: propertyPaneGroups
                }
            ]
        };
    }
}
