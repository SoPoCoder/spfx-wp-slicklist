import * as React from 'react';
import { Modal } from 'office-ui-fabric-react';
//import { IFieldInfo } from '@pnp/sp/fields';
import styles from './Slicklist.module.scss';
import { ISlickModalProps } from '../../..';
import { getFieldValue } from '../../../Utils';
//import linkifyHtml from 'linkify-html';

export default class SlickModal extends React.Component<ISlickModalProps> {
/*     private getItemFieldValue(item: IListItem, field: IFieldInfo): string {
        const strItem = item[field.InternalName];
        // if field value is a date, format it as a string
        if (field.TypeDisplayName === FieldTypes.DateTime) {
            return new Date(strItem).toLocaleDateString('en-US', { timeZone: 'UTC' });
        }
        // if field value is a file, format it as a link to the file
        if (field.TypeDisplayName === FieldTypes.File) {
            const fileLeafRef: string = item.FileLeafRef;
            const fileRef: string = item.FileRef;
            return `<a href='${fileRef}'>${fileLeafRef}</a>`;
        }
        // if field value is boolean, display yes or no
        if (field.TypeDisplayName === FieldTypes.Boolean)
            return item[field.InternalName] ? "Yes" : "No";
        // if field is a single line string, check if any hyperlinks are present and linkify them
        if (field.TypeDisplayName === FieldTypes.Single)
            return item[field.InternalName] ? linkifyHtml(item[field.InternalName], { defaultProtocol: "https" }) : "";
        // for all other field value types, simply display value as string
        return strItem;
    } */
    public render(): React.ReactElement<ISlickModalProps> {
        const { table1Fields, showModal, onClose, table1Item, table2Item } = this.props;
        const table2Fields = this.props.table2Fields ? this.props.table2Fields : [];
        return (
            <Modal isOpen={showModal} isBlocking={false} ignoreExternalFocusing={false} containerClassName={`${styles.slickmodal}`}>
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