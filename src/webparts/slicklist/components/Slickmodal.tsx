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
        const { table1Fields, showModal, onClose, table1Item, table2Item} = this.props;
        const table2Fields = this.props.table2Fields ? this.props.table2Fields : [];
        return (
            <Modal isOpen={showModal} isBlocking={false} containerClassName={`${styles.slickmodal}`}>
                <header>
                    <h2>{table1Fields[0] ? table1Fields[0].Title : "More"} Details</h2>
                    <button type="button" onClick={() => onClose(false)}>╳</button>
                </header>
                <table>
                    <tbody>{table1Fields.map((field, fieldIndex) => <tr key={fieldIndex}><th>{field.Description ? field.Description : field.Title}</th><td dangerouslySetInnerHTML={{ __html: this.getItemFieldValue(table1Item, field) }} /></tr>)}</tbody>
                    <tfoot>{table2Fields.map((field, fieldIndex) => <tr key={fieldIndex}><th>{field.Description ? field.Description : field.Title}</th><td dangerouslySetInnerHTML={{ __html: this.getItemFieldValue(table2Item, field) }} /></tr>)}</tfoot>
                </table>
            </Modal>
        );
    }
}