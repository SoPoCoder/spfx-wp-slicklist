import * as React from 'react';
import styles from './Slickmodal.module.scss';
import { Modal } from 'office-ui-fabric-react';
import { IFieldInfo } from '@pnp/sp/fields';
import { IListItem, ISlickModalProps } from '../../..';

export default class SlickModal extends React.Component<ISlickModalProps> {
    private getItemFieldValue(item: IListItem | undefined, field: IFieldInfo): string {
        if (item) {
            if (field.TypeDisplayName === "Yes/No")
                return item[field.InternalName] ? "Yes" : "No";
            return item[field.InternalName] ?? "";
        }
        return "";
    }
    public render(): React.ReactElement<ISlickModalProps> {
        return (
            <Modal isOpen={this.props.showModal} isBlocking={false} containerClassName={`${styles.slickmodal}`}>
                <div>
                    <h2>{this.props.fields[0] ? this.props.fields[0].Title : "More"} Details</h2>
                    <button type="button" onClick={() => this.props.onClose(false)}>╳</button>
                </div>
                <table>
                    {this.props.fields.map((field, fieldIndex) => <tr key={fieldIndex}><td>{field.Description ? field.Description : field.Title}</td><td>{this.getItemFieldValue(this.props.item, field)}</td></tr>)}
                </table>
            </Modal>
        );
    }
}