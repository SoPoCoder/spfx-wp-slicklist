import { IFieldInfo } from "@pnp/sp/fields";
import { FieldTypes, IListItem } from ".";

/* -----------------------------------------------------------------
    filters an array of IListItem taking into account field type
----------------------------------------------------------------- */
export function filterItems(filterField: IFieldInfo | undefined, filterValue: string, items: Array<IListItem>): Array<IListItem> {
    let filteredItems = new Array<IListItem>;
    if (filterField && filterValue) {
        filterValue = filterValue.toLowerCase();
        filteredItems = items.filter(item => {
            if (filterField) {
                const itemValue = item[filterField.InternalName];
                if (filterField.TypeDisplayName === FieldTypes.Boolean) { // boolean types
                    return itemValue.toString() === filterValue;
                }
                if (itemValue && filterField.TypeDisplayName === FieldTypes.DateTime) { // date types
                    const itemValueFormatted = new Date(itemValue).toLocaleDateString('en-US', { timeZone: 'UTC' });
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
    gets a table cells class based on column index and field type
----------------------------------------------------------------- */
export function getColumnClass(fieldType: string, fieldIndex: number, tableVisColsMobile: number, tableVisColsTablet: number, tableVisColsDesktop: number, lookupColIndex?: number): string | undefined {
    let className = undefined;
    lookupColIndex = lookupColIndex || 0;
    if (fieldIndex === 0 || fieldIndex === lookupColIndex) {
        className = "pcursor"
    }
    if (fieldType === FieldTypes.Boolean) {
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
    //skip for multi-line fields which may include lengthy hidden markup
    if (field.TypeDisplayName !== FieldTypes.Multiple) {

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
export function getFieldValue(item: IListItem, field: IFieldInfo): string {
    // if field value is a date, format it as a string
    if (field.TypeDisplayName === FieldTypes.DateTime) {
        return new Date(item[field.InternalName]).toLocaleDateString('en-US', { timeZone: 'UTC' });
    }
    // if field value is a Yes/No boolean, display checkmark for True and nothing for false
    if (field.TypeDisplayName === FieldTypes.Boolean) {
        return item[field.InternalName] ? "✓" : "";
    }
    // for all other field value types, simply display value as string
    return item[field.InternalName];
}