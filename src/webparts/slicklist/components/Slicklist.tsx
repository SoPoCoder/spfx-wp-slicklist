import * as React from 'react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { SPFI } from "@pnp/sp";
import { getSP } from '../../../utils/pnpjs-config';
import { IFieldInfo } from '@pnp/sp/fields';
import "@pnp/sp/lists";
import { FieldTypes, IListItem, ISlickListProps, ISlickListState } from '../../..';
import Table1 from './Table1';
import { Web } from '@pnp/sp/webs';
import styles from './Slicklist.module.scss';
import SlickModal from './SlickModal';
import Table2 from './Table2';

export default class Slicklist extends React.Component<ISlickListProps, ISlickListState> {

    private _sp: SPFI;

    constructor(props: ISlickListProps) {
        super(props);
        this._sp = getSP(this.context);
        this.state = {
            table1Fields: new Array<IFieldInfo>,
            table1Items: new Array<IListItem>,
            clickedItem: undefined,
            clickedLookup: "",
            table2Fields: new Array<IFieldInfo>,
            table2Items: new Array<IListItem>,
        };
        this.getListData(this.props.table1SiteURL, this.props.table1ListName, 1).catch((error: Error) => { throw error });
        this.getListData(this.props.table2SiteURL, this.props.table2ListName, 2).catch((error: Error) => { throw error });
    }

    /* -----------------------------------------------------------------
        gets fields/items from a SharePoint list based on url & name
    ----------------------------------------------------------------- */
    private async getListData(siteURL: string, listName: string, tableNumber?: number): Promise<void> {
        if (siteURL && listName) {
            const { orderByColumn1, orderByColumn2, orderByColumn3 } = this.props
            const web = Web([this._sp.web, siteURL]);
            const listFields: Array<IFieldInfo> = [];
            const listItems: Array<IListItem> = [];
            await web.lists.getByTitle(listName).fields.filter("ReadOnlyField eq false and Hidden eq false")().then((fields) => {
                if (fields) {
                    // get all the non-hidden fields of the following types
                    fields.map((field: IFieldInfo) => {
                        if (
                            field.TypeDisplayName === FieldTypes.Single ||
                            field.TypeDisplayName === FieldTypes.Multiple ||
                            field.TypeDisplayName === FieldTypes.Choice ||
                            field.TypeDisplayName === FieldTypes.Boolean ||
                            field.TypeDisplayName === FieldTypes.Number ||
                            field.TypeDisplayName === FieldTypes.DateTime) {
                            listFields.push(field);
                        }
                    });
                    // get all list items for Table1
                    if (tableNumber === 1)
                        web.lists.getByTitle(listName).items().then((items) => {
                            if (items) {
                                items.map((item) => {
                                    listItems.push(item);
                                })
                                this.setState({ table1Fields: listFields, table1Items: listItems });
                            }
                        }).catch((error: Error) => { throw error });
                    // get all list items for Table2
                    if (tableNumber === 2) {
                        let items = web.lists.getByTitle(listName).items;
                        items = orderByColumn1 ? items.orderBy(orderByColumn1) : items;
                        items = orderByColumn2 ? items.orderBy(orderByColumn2) : items;
                        items = orderByColumn3 ? items.orderBy(orderByColumn3, false) : items;
                        items.getAll().then((result) => {
                            result.map((item) => {
                                listItems.push(item);
                            })
                            this.setState({ table2Fields: listFields, table2Items: listItems });
                        }).catch((error: Error) => { throw error });
                    }
                }
            }).catch((error: Error) => { throw error });
        }
    }

    public componentDidUpdate(prevProps: ISlickListProps): void {
        // check to see if Table1 properties changed and update if so
        if (
            prevProps.table1SiteURL !== this.props.table1SiteURL ||
            prevProps.table1ListName !== this.props.table1ListName
        ) {
            this.getListData(this.props.table1SiteURL, this.props.table1ListName, 1).catch((error: Error) => { throw error });
        }

        // check to see if Table2 properties changed and update if so
        if (
            prevProps.table2SiteURL !== this.props.table2SiteURL ||
            prevProps.table2ListName !== this.props.table2ListName ||
            prevProps.orderByColumn1 !== this.props.orderByColumn1 ||
            prevProps.orderByColumn2 !== this.props.orderByColumn2 ||
            prevProps.orderByColumn3 !== this.props.orderByColumn3
        ) {
            this.getListData(this.props.table2SiteURL, this.props.table2ListName, 2).catch((error: Error) => { throw error });
        }
    }

    public render(): React.ReactElement {

        const placeholder = <Placeholder
            iconName="TableComputed"
            iconText="Configure your web part"
            description="Select a list to have it's contents rendered as a highly responsive filterable table."
            buttonLabel="Choose a List"
            onConfigure={this.props.onConfigure}
        />
        const Table1Frag = <Table1
            tableTitle={this.props.table1Title}
            tableVisColsMobile={this.props.table1VisColsMobile}
            tableVisColsTablet={this.props.table1VisColsTablet}
            tableVisColsDesktop={this.props.table1VisColsDesktop}
            fields={this.state.table1Fields}
            items={this.state.table1Items}
            lookupColumn={this.props.lookupColumn}
            onTopClick={this.props.onTopClick}
            onModalClick={(item: IListItem) => { this.setState({ clickedItem: item }) }}
            onLookupClick={this.props.onLookupClick}
        />
        const Table2Frag = <Table2
            tableTitle={this.props.table2Title}
            tableVisColsMobile={this.props.table2VisColsMobile}
            tableVisColsTablet={this.props.table2VisColsTablet}
            tableVisColsDesktop={this.props.table2VisColsDesktop}
            fields={this.state.table2Fields}
            items={this.state.table2Items}
            orderByColumn1={this.props.orderByColumn1}
            orderByColumn2={this.props.orderByColumn2}
            orderByColumn3={this.props.orderByColumn3}
            onTopClick={this.props.onTopClick}
        />
        const Modal = <SlickModal
            fields={this.state.table1Fields}
            item={this.state.clickedItem}
            showModal={this.state.clickedItem ? true : false}
            onClose={() => { this.setState({ clickedItem: undefined }) }}
        />

        // if only table1 list has been selected render Table1 and Modal, if lookup field has been selected, render Table1, Table2 and Modal, else show Placeholder
        if (this.props.table1ListName) {
            if (this.props.lookupColumn)
                return (<div className={`${styles.slicklist}`}>{React.createElement(React.Fragment, this.props, [Table1Frag, Table2Frag, Modal])}</div>);
            return (<div className={`${styles.slicklist}`}>{React.createElement(React.Fragment, this.props, [Table1Frag, Modal])}</div>);
        }
        return placeholder;
    }
}
