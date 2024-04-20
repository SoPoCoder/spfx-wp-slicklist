import * as React from 'react';
import "@pnp/sp/lists";
import { IListItem, ITable2Props, ITable2State } from '../../..';
import { getColumnClass, getFieldTitle, getFieldValue } from '../../../Utils';
import styles from './Slicklist.module.scss';
import { IFieldInfo } from '@pnp/sp/fields';

export default class Table2 extends React.Component<ITable2Props, ITable2State> {

    constructor(props: ITable2Props) {
        super(props);
        this.state = {
            fields: new Array<IFieldInfo>,
            items: new Array<IListItem>
        };
    }

    public componentDidUpdate(prevProps: ITable2Props): void {
        // set state of fields and items to passed in properties when component is instantiated
        if (
            prevProps.fields !== this.props.fields ||
            prevProps.items !== this.props.items
        ) {
            this.setState({
                fields: this.props.fields,
                items: this.props.items
            })
        }
    }

    public render(): React.ReactElement<ITable2Props> {
        const { tableTitle, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop, orderByColumn1, orderByColumn3 } = this.props;
        const { fields, items } = this.state;

        let currentUnit: string = "";
        const getHeaderRows = (unit: string): React.ReactFragment => {
            if (unit && unit !== currentUnit) {
                currentUnit = unit;
                return (
                    <>
                        <tr className={`${styles.title}`}><th colSpan={fields.length}><span className={`${styles.totop} pcursor`} onClick={() => this.props.onTopClick()}>&#9650; TOP</span>{unit}</th></tr>
                        <tr>{fields.map((field, fieldIndex) => <th className={getColumnClass(field.TypeDisplayName, fieldIndex, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop)} key={fieldIndex} title={field.Description}>{getFieldTitle(field, items)}</th>)}</tr>
                    </>
                )
            }
            return (<></>);
        }
        return (
            <table className={`${styles.buffer}`}>
                <tbody>
                    {items.map((item, itemIndex) =>
                        <>
                            {tableTitle && itemIndex === 0 ? <tr className={`${styles.title}`}><th colSpan={fields.length}>{tableTitle}</th></tr> : getHeaderRows(item[orderByColumn1])}
                            <tr key={itemIndex} id={item.Title} className={orderByColumn3 && item[orderByColumn3] ? `${styles.grouping2}` : undefined}>{
                                fields.map((field, fieldIndex) => <td className={getColumnClass(field.TypeDisplayName, fieldIndex, tableVisColsMobile, tableVisColsTablet, tableVisColsDesktop)} key={fieldIndex} dangerouslySetInnerHTML={{ __html: getFieldValue(item, field) }} />)}
                            </tr>
                        </>
                    )}
                </tbody>
                <tfoot>
                    <tr><th colSpan={fields.length}><span className={`${styles.totop} pcursor`} onClick={() => this.props.onTopClick()}>&#9650; TOP</span>&nbsp;</th></tr>
                </tfoot>
            </table>
        );
    }

}