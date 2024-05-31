import * as React from 'react';
import { Modal } from 'office-ui-fabric-react';
import styles from './Slicklist.module.scss';
import { ISlickModalProps } from '../../..';
import { getFieldValue } from '../../../Utils';

export default class SlickModal extends React.Component<ISlickModalProps> {
    public render(): React.ReactElement<ISlickModalProps> {
        const { table1Fields, showModal, onClose, table1Item, table2Item } = this.props;
        let table2Fields = this.props.table2Fields ? this.props.table2Fields : [];
        if (table2Fields.length > 0 && this.props.orderByFields) {
            table2Fields = table2Fields.filter((field) => this.props.orderByFields?.indexOf(field.InternalName) === -1)
        }
        return (
            <Modal isOpen={showModal} isBlocking={true} ignoreExternalFocusing={true} containerClassName={`${styles.slickmodal}`}>
                <header>
                    <h2>{table1Fields[0] ? table1Fields[0].Title : "More"} Details</h2>
                    <button type="button" onClick={() => onClose(false)}>╳</button>
                </header>
                <table>
                    <tbody>{table1Fields.map((field, fieldIndex) => <tr key={fieldIndex}><th title={field.Description}>{field.Title}</th><td dangerouslySetInnerHTML={{ __html: getFieldValue(table1Item, field, true) }} /></tr>)}</tbody>
                    <tfoot>{table2Fields.map((field, fieldIndex) => <tr key={fieldIndex}><th title={field.Description}>{field.Title}</th><td dangerouslySetInnerHTML={{ __html: getFieldValue(table2Item, field, true) }} /></tr>)}</tfoot>
                </table>
            </Modal>
        );
    }
}