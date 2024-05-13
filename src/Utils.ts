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
                const itemValue = item[filterField.InternalName] ? item[filterField.InternalName].toString() : "";
                if (filterField.TypeDisplayName === FieldTypes.Boolean) { // boolean types
                    return itemValue === filterValue;
                }
                if (itemValue && filterField.TypeDisplayName === FieldTypes.DateTime) { // date types
                    const itemValueFormatted = new Date(itemValue).toLocaleDateString('en-US', { timeZone: 'UTC' });
                    const selectedValueFormatted = new Date(filterValue).toLocaleDateString('en-US', { timeZone: 'UTC' });
                    return itemValueFormatted === selectedValueFormatted;
                }
                if (itemValue) { // for all other types
                    const fieldValue = itemValue.toLowerCase();
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
export function getColumnClass(isRow: boolean, fieldType: string, fieldIndex: number, tableVisColsMobile: number, tableVisColsTablet: number, tableVisColsDesktop: number, lookupColIndex?: number): string | undefined {
    let className = undefined;
    if (isRow) {
        lookupColIndex = lookupColIndex || 0;
        if (fieldIndex === 0 || fieldIndex === lookupColIndex) {
            className = "pcursor"
        }
        if (fieldType === FieldTypes.File || fieldType === FieldTypes.Boolean) {
            className = className ? className.concat(" ", "mark") : "mark";
        }
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
    //skip for multi-line fields which may include lengthy hidden markup (recommend leave these columns at end of table as their width will change with filtering)
    if (field.TypeDisplayName !== FieldTypes.File && field.TypeDisplayName !== FieldTypes.Multiple) {

        //get the length in characters of the longest entry for the given field
        let longestFieldValue = 0;
        items.map((item: IListItem) => {
            if (item[field.InternalName])
                if (item[field.InternalName].toString().length > longestFieldValue)
                    longestFieldValue = item[field.InternalName].toString().length;
        });

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
    let strItem: string = "";
    if (item && field && item[field.InternalName]) {
        // get value for specified item and field
        strItem = item[field.InternalName].toString();
        // if field value is a file, format it as a link to the file
        if (field.TypeDisplayName === FieldTypes.File) {
            const fileLeafRef: string = item.FileLeafRef.toString();
            const fileRef: string = item.FileRef.toString();
            if (isModalTable)
                return `<a href="${fileRef}" target="_blank" data-interception="off">${fileLeafRef}</a>`;
            else
                return `<i class="${getIconClassName(getFileIcon(fileLeafRef))}" title="${fileLeafRef}" />`;
        }
        // if field value is a date, format it as a string
        if (field.TypeDisplayName === FieldTypes.DateTime) {
            return new Date(strItem).toLocaleDateString('en-US', { timeZone: 'UTC' });
        }
        // if field value is a Yes/No boolean, display checkmark for True and nothing for false
        if (field.TypeDisplayName === FieldTypes.Boolean) {
            if (isModalTable)
                return strItem ? "Yes" : "No";
            else
                return strItem ? "✓" : "";
        }
        // if field is a single line string, check if any hyperlinks are present and linkify them
        if (field.TypeDisplayName === FieldTypes.Single) {
            return strItem ? linkifyHtml(strItem, { defaultProtocol: "https", target: "_blank" }) : "";
        }
        // if field is a hyperlink, display as a hyperlink
        if (field.TypeDisplayName === FieldTypes.Link) {
            const hyperLink = item[field.InternalName] as HyperLink;
            return hyperLink ? `<a href="${hyperLink.Url}" target="_blank" data-interception="off">${hyperLink.Description}</a>` : "";
        }
    }
    // for all other field value types, simply display value as string
    return strItem;
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
    }
    return "Document";
}