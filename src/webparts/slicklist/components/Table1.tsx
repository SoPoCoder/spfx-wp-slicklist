import * as React from 'react';
import "@pnp/sp/lists";
import { FieldTypes, IListItem, ITable1Props, ITable1State } from '../../..';
import { filterItems, getColumnClass, getFieldTitle, getFieldValue } from '../../../Utils';
import styles from './Slicklist.module.scss';
import { IFieldInfo } from '@pnp/sp/fields';

export default class Table1 extends React.Component<ITable1Props, ITable1State> {

    constructor(props: ITable1Props) {
        super(props);
        this.state = {
            fields: new Array<IFieldInfo>,
            items: new Array<IListItem>,
            filterField: undefined,
            filterValue: "",
        };
    }

    public componentDidUpdate(prevProps: ITable1Props, prevState: ITable1State): void {
        // set state of fields and items to passed in properties when component is instantiated
        if (
            prevProps.fields !== this.props.fields ||
            prevProps.items !== this.props.items
        ) {
            this.setState({
                fields: this.props.fields,
                // items: this.props.items // uncomment to show rather than hide all items on load
            })
        }
        //if a field and value was selected, filter items and update state
        if (
            prevState.filterField !== this.state.filterField ||
            prevState.filterValue !== this.state.filterValue
        ) {
            this.setState({ items: filterItems(this.state.filterField, this.state.filterValue, this.props.items) });
        }
    }

    /* -----------------------------------------------------------------
        gets an input element based on a given columns field type
    ----------------------------------------------------------------- */
    private getFieldInput(field: IFieldInfo, filterField: IFieldInfo | undefined, filterValue: string): React.ReactElement {
        const fieldValue: string = (filterField && field.InternalName === filterField.InternalName) ? filterValue : "";

        // return a date type input when the column field type is of type DateTime
        if (field.TypeDisplayName === FieldTypes.DateTime) {
            return <input type="date" id={field.InternalName} name={field.InternalName} placeholder={"[ Enter " + field.Title + " ]"} onChange={(e) => { this.setState({ filterField: field, filterValue: e.target.value }) }} value={fieldValue} />
        }

        // return a select menu with Yes, No and no selection (default) if column type is of type Yes/No (Boolean)
        if (field.TypeDisplayName === FieldTypes.Boolean) {
            return <select id={field.InternalName} name={field.InternalName} onChange={(e) => { this.setState({ filterField: field, filterValue: e.target.value }) }} value={fieldValue} ><option value="">―</option><option value='true'>Yes</option><option value='false'>No</option></select>
        }

        // return a select menu with unique choices for that column plus no selection (default) if column is of type Choice
        if (field.TypeDisplayName === FieldTypes.Choice) {
            let choices: Array<string> = [];
            this.props.items.map((item: IListItem) => {
                const newChoice: string | null = item[field.InternalName] ? item[field.InternalName] : null;
                if (newChoice && choices.indexOf(newChoice) < 0)
                    choices.push(newChoice);
            });
            choices = choices.sort((x, y) => x > y ? 1 : x < y ? -1 : 0);
            return (
                <select id={field.InternalName} name={field.InternalName} value={fieldValue} onChange={(e) => { this.setState({ filterField: field, filterValue: e.target.value }) }} >
                    <option value="">―</option>
                    {choices.map((choice) =>
                        <option key={choice} value={choice}>{choice}</option>
                    )}
                </select>
            );
        }

        // return a text type input in all other cases
        return <input type="text" id={field.InternalName} name={field.InternalName} placeholder={"[ Enter " + field.Title + " ]"} onChange={(e) => { this.setState({ filterField: field, filterValue: e.target.value }) }} value={fieldValue} />
    }

    /* -----------------------------------------------------------------
        passes click events up to parent component based on field index
    ----------------------------------------------------------------- */
    private onClickHandler(event: React.MouseEvent, fieldIndex: number, selectedItem: IListItem, lookupColIndex: number): void {
        if (fieldIndex === 0)
            this.props.onModalClick(selectedItem);
        else
            if (fieldIndex === lookupColIndex)
                this.props.onLookupClick(event.currentTarget.innerHTML)
    }

    public render(): React.ReactElement<ITable1Props> {
        const { tableTitle, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop, lookupColumn } = this.props;
        const { fields, items, filterField, filterValue } = this.state;
        const lookupColIndex = this.props.fields.map(item => item.InternalName).indexOf(lookupColumn);
        return (
            <table>
                <thead>
                    <tr className={`${styles.title}`}><th colSpan={fields.length}>{tableTitle}</th></tr>
                    <tr>
                        {fields.map((field, fieldIndex) => <th className={getColumnClass(field.TypeDisplayName, fieldIndex, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop)} key={fieldIndex} title={field.Description}>{getFieldTitle(field, this.props.items)}</th>)}
                    </tr>
                    <tr className={`${styles.fields}`} id="fields">
                        {fields.map((field, fieldIndex) => <th className={getColumnClass(field.TypeDisplayName, fieldIndex, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop)} key={fieldIndex}>{this.getFieldInput(field, filterField, filterValue)}</th>)}
                    </tr>
                </thead>
                <tbody>
                    {items.map((item, itemIndex) => <tr key={itemIndex} id={item.Title}>{
                        fields.map((field, fieldIndex) => <td className={getColumnClass(field.TypeDisplayName, fieldIndex, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop, lookupColIndex)} key={fieldIndex} onClick={(e) => this.onClickHandler(e, fieldIndex, item, lookupColIndex)} dangerouslySetInnerHTML={{ __html: getFieldValue(item, field) }} />)
                    }</tr>)}
                </tbody>
            </table>
        );
    }

}
