// A file is required to be in the root of the /src directory by the TypeScript compiler
import { IFieldInfo } from "@pnp/sp/fields";

export interface ISPSite {
    SPSiteUrl: string; // do not rename
    Title: string; // do not rename
}

export interface ISPLists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
}

export interface IListItem {
    [index: string]: string;
}

export interface ITable {
    tableTitle: string;
    tableVisColsMobile: number;
    tableVisColsTablet: number;
    tableVisColsDesktop: number;
    fields: Array<IFieldInfo>;
    items: Array<IListItem>;
    onTopClick: () => void;
}

export enum FieldTypes {
    Single = "Single line of text",
    Multiple = "Multiple lines of text",
    Choice = "Choice",
    Boolean = "Yes/No",
    Number = "Number",
    DateTime = "Date and Time"
}

/*-----------------------------------------------------
SlicklistWebPart Interfaces
-----------------------------------------------------*/
export interface ISlicklistWebPartProps {
    // table 1 properties
    table1Title: string;
    table1SiteURL: string;
    table1ListName: string;
    table1VisColsMobile: number;
    table1VisColsTablet: number;
    table1VisColsDesktop: number;
    // table 2 properties
    lookupColumn: string;
    table2Title: string;
    table2SiteURL: string;
    table2ListName: string;
    table2VisColsMobile: number;
    table2VisColsTablet: number;
    table2VisColsDesktop: number;
    orderByColumn1: string;
    orderByColumn2: string;
    orderByColumn3: string;
}

/*-----------------------------------------------------
SlickList Interfaces
-----------------------------------------------------*/
export interface ISlickListProps {
    // table 1 properties
    table1Title: string;
    table1SiteURL: string;
    table1ListName: string;
    table1VisColsMobile: number;
    table1VisColsTablet: number;
    table1VisColsDesktop: number;
    // table 2 properties
    lookupColumn: string;
    table2Title: string;
    table2SiteURL: string;
    table2ListName: string;
    table2VisColsMobile: number;
    table2VisColsTablet: number;
    table2VisColsDesktop: number;
    orderByColumn1: string;
    orderByColumn2: string;
    orderByColumn3: string;
    // callback functions
    onConfigure: () => void;
    onTopClick: () => void;
    onLookupClick: (value: string) => void;
}

export interface ISlickListState {
    // table 1 state
    table1Fields: Array<IFieldInfo>;
    table1Items: Array<IListItem>;
    clickedItem: IListItem | undefined;
    clickedLookup: string;
    // table 2 state
    table2Fields: Array<IFieldInfo>;
    table2Items: Array<IListItem>;
}

/*-----------------------------------------------------
Table1 Interfaces
-----------------------------------------------------*/
export interface ITable1Props extends ITable {
    lookupColumn: string;
    onModalClick: (item: IListItem) => void;
    onLookupClick: (value: string) => void;
}

export interface ITable1State {
    fields: Array<IFieldInfo>;
    items: Array<IListItem>;
    filterField: IFieldInfo | undefined;
    filterValue: string;
}

/*-----------------------------------------------------
Table2 Interfaces
-----------------------------------------------------*/
export interface ITable2Props extends ITable {
    orderByColumn1: string;
    orderByColumn2: string;
    orderByColumn3: string;
    onTopClick: () => void;
}

export interface ITable2State {
    fields: Array<IFieldInfo>;
    items: Array<IListItem>;
}

/*-----------------------------------------------------
SlickModal Interfaces
-----------------------------------------------------*/
export interface ISlickModalProps {
    table1Fields: Array<IFieldInfo>;
    table1Item: IListItem | undefined;
    table2Fields: Array<IFieldInfo>;
    table2Item: IListItem | undefined;
    showModal: boolean;
    onClose: (value: boolean) => void;
}