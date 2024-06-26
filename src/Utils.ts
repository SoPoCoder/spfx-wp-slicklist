import { IFieldInfo } from "@pnp/sp/fields";
import { FieldTypes, HyperLink, IListItem } from ".";
import linkifyHtml from "linkify-html";
import { getIconClassName } from "office-ui-fabric-react";

/* -----------------------------------------------------------------
    filters an array of IListItem taking into account field type
----------------------------------------------------------------- */
export function filterItems(filterField: IFieldInfo | undefined, filterValue: string, items: Array<IListItem>): Array<IListItem> {
    let filteredItems = new Array<IListItem>;
    if (filterField && filterValue) {
        filterValue = filterValue.toLowerCase();
        filteredItems = items.filter(item => {
            if (filterField) {
                let itemValue = item[filterField.InternalName];
                if (filterField.TypeDisplayName === FieldTypes.Boolean) { // boolean types
                    itemValue = itemValue ? "true" : "false";
                    return itemValue === filterValue;
                }
                if (itemValue && filterField.TypeDisplayName === FieldTypes.DateTime) { // date types
                    const itemValueFormatted = new Date(itemValue.toString()).toLocaleDateString('en-US', { timeZone: 'UTC' });
                    const selectedValueFormatted = new Date(filterValue).toLocaleDateString('en-US', { timeZone: 'UTC' });
                    return itemValueFormatted === selectedValueFormatted;
                }
                if (itemValue) { // for all other types
                    const fieldValue = itemValue.toString().toLowerCase();
                    const start1Found = fieldValue.indexOf(filterValue) === 0;
                    const start2Found = fieldValue.indexOf(" " + filterValue) > -1;
                    return start1Found || start2Found;
                }
                return false;
            }
        });
    }
    return filteredItems;
}

/* -----------------------------------------------------------------
gets item from table2Items based on selected item from table1Items
----------------------------------------------------------------- */
export function getTable2Item(lookupColumn: string | undefined, table1Item: IListItem | undefined, table2Items: IListItem[]): IListItem | undefined {
    // gets the value of the defined lookup column for the user selected item from table1Items
    const clickedLookupColValue = lookupColumn && table1Item ? table1Item[lookupColumn] : "";
    // filters table2Items on the defined lookup column based on the value above
    const table2ItemsFiltered: IListItem[] | undefined = table2Items.filter(item => { return item.Title === clickedLookupColValue });
    // returns the first entry of the filtered list if items exist, else returns undefined
    return table2ItemsFiltered && table2ItemsFiltered.length > -1 ? table2ItemsFiltered[0] : undefined;
}

/* -----------------------------------------------------------------
    gets a table cells class based on column index and field type
----------------------------------------------------------------- */
export function getColumnClass(isRow: boolean, field: IFieldInfo, fieldIndex: number, tableVisColsMobile: number, tableVisColsTablet: number, tableVisColsDesktop: number, lookupColIndex?: number): string | undefined {
    let className = undefined;
    if (isRow) {
        lookupColIndex = lookupColIndex || 0;
        if (field.InternalName === "Title" || fieldIndex === lookupColIndex) {
            className = "pcursor"
        }
    }
    if (field.TypeDisplayName === FieldTypes.File || field.TypeDisplayName === FieldTypes.Boolean) {
        className = className ? className.concat(" ", "mark") : "mark";
    }
    if (fieldIndex >= tableVisColsMobile) {
        if (fieldIndex >= tableVisColsTablet) {
            if (fieldIndex >= tableVisColsDesktop) {
                return className ? className.concat(" ", "hideFromDesktop") : "hideFromDesktop";
            }
            return className ? className.concat(" ", "hideFromTablet") : "hideFromTablet";
        }
        return className ? className.concat(" ", "hideFromMobile") : "hideFromMobile";
    }
    return className;
}

/* -----------------------------------------------------------------
    gets a columns title and formats it based on longest value
----------------------------------------------------------------- */
export function getFieldTitle(field: IFieldInfo, items: Array<IListItem>): string {
    if (field.TypeDisplayName !== FieldTypes.File) {

        //get the length in characters of the longest entry for the given field
        let longestFieldValue = 0;
        if (field.TypeDisplayName === FieldTypes.Multiple) {
            longestFieldValue = 34; // set Multi-line text column with to 34 since that is truncated to 32
        } else {
            items.map((item: IListItem) => {
                if (item[field.InternalName])
                    if (item[field.InternalName].toString().length > longestFieldValue)
                        longestFieldValue = item[field.InternalName].toString().length;
            });
        }

        //based on the longest value and the length of the field title, calculate how many spaces are required on both sides
        if (longestFieldValue > field.Title.length) {
            const spaceBufferLength = Math.ceil((longestFieldValue - field.Title.length) / 2);
            let spaceBuffer = "";
            for (let i = 0; i < spaceBufferLength; i++) {
                spaceBuffer = spaceBuffer + "\u00A0 "; //note that this adds two spaces, one non-breaking and one normal
            }
            return (spaceBuffer + field.Title + spaceBuffer);
        }
    }
    return field.Title;
}

/* -----------------------------------------------------------------
    gets a table cells value and formats it based on field type
----------------------------------------------------------------- */
export function getFieldValue(item: IListItem | undefined, field: IFieldInfo, isModalTable: boolean = false): string {
    let strItem = undefined;
    if (item && field) {
        // get value for specified item and field
        strItem = item[field.InternalName];
        // if field value is a file, format it as a link to the file
        if (field.TypeDisplayName === FieldTypes.File) {
            const fileLeafRef: string = item.FileLeafRef.toString();
            const fileRef: string = item.FileRef.toString();
            if (isModalTable)
                return `<a href="${fileRef}" target="_blank" data-interception="off">${fileLeafRef}</a>`;
            else
                return `<a href="${fileRef}" target="_blank" data-interception="off"><i class="${getIconClassName(getFileIcon(fileLeafRef))}" title="${fileLeafRef}" /></a>`;
        }
        // if field value is a date, format it as a string
        if (field.TypeDisplayName === FieldTypes.DateTime && strItem) {
            return new Date(strItem.toString()).toLocaleDateString('en-US', { timeZone: 'UTC' });
        }
        // if field value is a Yes/No boolean, display checkmark for True and nothing for false
        if (field.TypeDisplayName === FieldTypes.Boolean) {
            if (isModalTable)
                return strItem ? "Yes" : "No";
            else
                return strItem ? "✓" : "";
        }
        // if field is a single line string, check if any hyperlinks are present and linkify them
        if (field.TypeDisplayName === FieldTypes.Single && strItem) {
            return strItem ? linkifyHtml(strItem.toString(), { defaultProtocol: "https", target: "_blank" }) : "";
        }
        // if field is a multi-line string, check if length is more than 30 characters and truncate with tooltip for list only
        if (field.TypeDisplayName === FieldTypes.Multiple && !isModalTable && strItem) {
            if (strItem.toString().substring(0,1) !== "<" && strItem.toString().length > 30)
                return `<span title="${strItem}">${strItem.toString().substring(0, 29) + "..."}</span>`;
        }
        // if field is a hyperlink, display as a hyperlink
        if (field.TypeDisplayName === FieldTypes.Link) {
            const hyperLink = item[field.InternalName] as HyperLink;
            return hyperLink ? `<a href="${hyperLink.Url}" target="_blank" data-interception="off">${hyperLink.Description}</a>` : "";
        }
    }
    // for all other field value types, simply display value as string
    return strItem ? strItem.toString() : "";
}

/* -----------------------------------------------------------------
    recieves a filename and returns an Office UI Fabric Icon name
----------------------------------------------------------------- */
export function getFileIcon(fileName: string | undefined): string {
    if (fileName) {
        const fileExt = fileName.split('.').pop();
        if (fileExt === "pdf")
            return "PDF";
        if (fileExt === "doc" || fileExt === "docx")
            return "WordDocument";
        if (fileExt === "xls" || fileExt === "xlsx")
            return "ExcelDocument";
        if (fileExt === "ppt" || fileExt === "pptx")
            return "PowerPointDocument";
        if (fileExt === "txt")
            return "TextDocument";
        if (fileExt === "png" || fileExt === "jpg" || fileExt === "gif" || fileExt === "jpeg")
            return "PictureFill";
        if (fileExt === "vsd" || fileExt === "vsdx")
            return "VisioDocument";
        if (fileExt === "exe" || fileExt === "msi")
            return "Product";
    }
    return "Document";
}