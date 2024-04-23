import * as React from 'react';
import { Modal } from 'office-ui-fabric-react';
import { IFieldInfo } from '@pnp/sp/fields';
import styles from './Slicklist.module.scss';
import { FieldTypes, IListItem, ISlickModalProps } from '../../..';

export default class SlickModal extends React.Component<ISlickModalProps> {
    private getItemFieldValue(item: IListItem | undefined, field: IFieldInfo): string {
        if (item) {
            if (field.TypeDisplayName === FieldTypes.Boolean)
                return item[field.InternalName] ? "Yes" : "No";
            return item[field.InternalName] ?? "";
        }
        return "";
    }
    public render(): React.ReactElement<ISlickModalProps> {
        // get the maximum field Title length in characters to calculate how wide to make the field Title column
/*         let fieldCount: number = 0;
        let maxFieldLen: number = 0;
        this.props.table1Fields.map(field => {
            fieldCount++;
            if (field.Title.length > maxFieldLen)
                maxFieldLen = field.Title.length;
        })
        this.props.table2Fields.map(field => {
            fieldCount++;
            if (field.Title.length > maxFieldLen)
                maxFieldLen = field.Title.length;
        }) */
        return (
            <Modal isOpen={this.props.showModal} isBlocking={false} containerClassName={`${styles.slickmodal}`}>
                <div>
                    <h2>{this.props.table1Fields[0] ? this.props.table1Fields[0].Title : "More"} Details</h2>
                    <button type="button" onClick={() => this.props.onClose(false)}>╳</button>
                </div>
                <table>
                    <tbody>{this.props.table1Fields.map((field, fieldIndex) => <tr key={fieldIndex}><th>{field.Description ? field.Description : field.Title}</th><td>{this.getItemFieldValue(this.props.table1Item, field)}</td></tr>)}</tbody>
                    <tfoot>{this.props.table2Fields.map((field, fieldIndex) => <tr key={fieldIndex}><th>{field.Description ? field.Description : field.Title}</th><td>{this.getItemFieldValue(this.props.table2Item, field)}</td></tr>)}</tfoot>
                </table>
            </Modal>
        );
    }
}