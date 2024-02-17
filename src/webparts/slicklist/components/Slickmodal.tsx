import * as React from 'react';
import styles from './Slickmodal.module.scss';
import { Modal } from 'office-ui-fabric-react';
import { IFieldInfo } from '@pnp/sp/fields';
import { IListItem } from './Slicklist';

export interface ISlickmodalProps {
    fields: Array<IFieldInfo>;
    item: IListItem | undefined;
    showModal: boolean;
    onClose: (value: boolean) => void;
}

export default class Slickmodal extends React.Component<ISlickmodalProps> {

    public render(): React.ReactElement<ISlickmodalProps> {
        return (
            <Modal isOpen={this.props.showModal} isBlocking={false} containerClassName={`${styles.slickmodal}`}>
                <div>
                    <h2>My Modal Popup</h2>
                    <button type="button" onClick={() => this.props.onClose(false)}>╳</button>
                </div>
                <table>
                    {this.props.fields.map((field, fieldIndex) => <tr key={fieldIndex}><td>{field.Title}</td><td>{this.props.item ? this.props.item[field.InternalName] : ""}</td></tr>)}
                </table>
            </Modal>
        );
    }
}