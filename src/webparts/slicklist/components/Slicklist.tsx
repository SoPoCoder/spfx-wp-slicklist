import * as React from 'react';
import styles from './Slicklist.module.scss';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { SPFI } from "@pnp/sp";
import { getSP } from '../../../utils/pnpjs-config';
import { Web } from '@pnp/sp/webs';
import { IFieldInfo } from '@pnp/sp/fields';
import "@pnp/sp/lists";
import { DisplayMode } from '@microsoft/sp-core-library';
import Slickmodal from './Slickmodal';
//import { escape } from '@microsoft/sp-lodash-subset';

export interface IListItem {
  [index: string]: string;
}

export interface ISlicklistProps {
  table1Title: string;
  table1SiteURL: string;
  table1ListName: string;
  visibleColsMobile: number;
  visibleColsTablet: number;
  visibleColsDesktop: number;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  onConfigure: () => void;
}

export interface ISlicklistState {
  fields: Array<IFieldInfo>;
  items: Array<IListItem>;
  filterField: string;
  filterValue: string;
  filteredItems: Array<IListItem>;
  selectedItem: IListItem | undefined;
  showModal: boolean;
}

export default class Slicklist extends React.Component<ISlicklistProps, ISlicklistState> {

  private _sp: SPFI;

  constructor(props: ISlicklistProps) {
    super(props);
    this.state = {
      fields: new Array<IFieldInfo>,
      items: new Array<IListItem>,
      filterField: "",
      filterValue: "",
      filteredItems: new Array<IListItem>,
      selectedItem: undefined,
      showModal: false
    };

    this._sp = getSP(this.context);
    this.getListItems().catch((error: Error) => { throw error });
  }

  public componentDidUpdate(prevProps: ISlicklistProps, prevState: ISlicklistState): void {
    // check to see if anything in properties panel changed and update webpart if so
    if (
      prevProps.table1Title !== this.props.table1Title ||
      prevProps.table1SiteURL !== this.props.table1SiteURL ||
      prevProps.table1ListName !== this.props.table1ListName ||
      prevProps.displayMode !== this.props.displayMode
    ) {
      this.getListItems().catch((error: Error) => { throw error });
    }
    // check to see if anything in the webpart changed and update webpart if so
    if (
      prevState.filterField !== this.state.filterField ||
      prevState.filterValue !== this.state.filterValue
    ) {

      //if a field and value was selected, reset fields, filter items and put into filterItems
      let filteredItems = new Array<IListItem>;
      if (this.state.filterField && this.state.filterValue) {
        const selectedField = this.state.filterField;
        const selectedValue = this.state.filterValue;
        filteredItems = this.state.items.filter(item => {
          if (item[selectedField]) {
            const fieldValue = item[selectedField].toString();
            const start1Found = fieldValue.toLowerCase().indexOf(selectedValue.toLowerCase()) === 0;
            const start2Found = fieldValue.toLowerCase().indexOf(" " + selectedValue.toLowerCase()) > -1;
            return start1Found || start2Found;
          }
        });
      }
      this.setState({ filteredItems: filteredItems });
    }
  }

  private async getListItems(): Promise<void> {
    // get a list of items from the specified list and set it to state
    if (typeof this.props.table1ListName !== "undefined" && this.props.table1ListName.length > 0) {
      const web = Web([this._sp.web, this.props.table1SiteURL]);

      const listFields: Array<IFieldInfo> = [];
      const listItems: Array<IListItem> = [];
      // get all the non-hidden fields for a list
      web.lists.getByTitle(this.props.table1ListName).fields.filter("ReadOnlyField eq false and Hidden eq false")().then((fields) => {
        if (fields) {
          fields.map((field: IFieldInfo) => {
            if (
              field.TypeDisplayName === "Text" ||
              field.TypeDisplayName === "Single line of text" ||
              field.TypeDisplayName === "Choice" ||
              field.TypeDisplayName === "Yes/No" ||
              field.TypeDisplayName === "Date and Time") {
              listFields.push(field);
            }
          })
          // get all items in the list
          web.lists.getByTitle(this.props.table1ListName).items().then((items) => {
            if (items) {
              items.map((item: IListItem) => {
                listItems.push(item);
              })
              this.setState({
                fields: listFields,
                items: listItems
              })
            }
          }).catch((error: Error) => { throw error });
        }
      }).catch((error: Error) => { throw error });

    }
  }

  private setColumnHideClass(className: string, fieldIndex: number): string {
    if (fieldIndex === 0)
      className = "pcursor"
    if (fieldIndex >= this.props.visibleColsMobile) {
      if (fieldIndex >= this.props.visibleColsTablet) {
        if (fieldIndex >= this.props.visibleColsDesktop) {
          return className ? className.concat(" ", "hideFromDesktop") : "hideFromDesktop";
        }
        return className ? className.concat(" ", "hideFromTablet") : "hideFromTablet";
      }
      return className ? className.concat(" ", "hideFromMobile") : "hideFromMobile";
    }
    return className;
  }

  private getFieldTitle(field: IFieldInfo): string {
    //get the length in characters of the longest entry for the given field
    let longestFieldValue = 0;
    this.state.items.map((item: IListItem) => {
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
    return field.Title;
  }

  private getFieldInput(field: IFieldInfo, filterField: string, filterValue: string): React.ReactElement {
    // determine the correct field input to choose based on the field type and set value to nothing, unless field internal name matches selected field in state
    function getFieldValue(fieldName: string): string {
      if (fieldName === filterField)
        return filterValue;
      return "";
    }
    if (field.TypeDisplayName === "Date and Time") {
      return <input type="date" id={field.InternalName} name={field.InternalName} placeholder={"[ Enter " + field.Title + " ]"} onChange={(e) => { this.setState({ filterField: field.InternalName, filterValue: e.target.value }) }} value={ getFieldValue(field.InternalName) } />
    }
    if (field.TypeDisplayName === "Yes/No") {
      return <select id={field.InternalName} name={field.InternalName} onChange={(e) => { this.setState({ filterField: field.InternalName, filterValue: e.target.value }) }} value={ getFieldValue(field.InternalName) } ><option value="">―</option><option value='Yes'>Yes</option><option value='No'>No</option></select>
    }
    if (field.TypeDisplayName === "Choice") {
      let choices: Array<string> = [];
      this.state.items.map((item: IListItem) => {
        const newChoice: string | null = item[field.InternalName] ? item[field.InternalName][0] : null;
        if (newChoice && choices.indexOf(newChoice) < 0)
          choices.push(newChoice);
      });
      choices = choices.sort((x, y) => x > y ? 1 : x < y ? -1 : 0);
      return (
        <select id={field.InternalName} name={field.InternalName} value={ getFieldValue(field.InternalName) } onChange={(e) => { this.setState({ filterField: field.InternalName, filterValue: e.target.value }) }} >
          <option value="">―</option>
          {choices.map((choice) =>
            <option key={choice} value={choice}>{choice}</option>
          )}
        </select>
      );
    }
    return <input type="text" id={field.InternalName} name={field.InternalName} placeholder={"[ Enter " + field.Title + " ]"} onChange={(e) => { this.setState({ filterField: field.InternalName, filterValue: e.target.value }) }} value={ getFieldValue(field.InternalName) } />
  }

  private onClickHandler(fieldIndex: number, selectedItem: IListItem): void {
    if (fieldIndex < 1)
      this.setState({ showModal: true, selectedItem: selectedItem })
  }

  public render(): React.ReactElement<ISlicklistProps> {
    //render the webpart when state changes
    const listSelected: boolean = typeof this.props.table1ListName !== "undefined" && this.props.table1ListName.length > 0;
    const { fields, filteredItems, filterField, filterValue } = this.state;
    return (
      <section>
        {!listSelected && (
          <Placeholder
            iconName="TableComputed"
            iconText="Configure your web part"
            description="Select a list to have it's contents rendered as a highly responsive filterable table."
            buttonLabel="Choose a List"
            onConfigure={this.props.onConfigure}
          />
        )}
        {listSelected && (
          <div className={`${styles.slicklist}`}>
            <table>
              <thead>
                <tr className={`${styles.title}`}><th colSpan={fields.length}>{this.props.table1Title}</th></tr>
                <tr>
                  {fields.map((field, fieldIndex) => <th className={this.setColumnHideClass("", fieldIndex)} key={fieldIndex}>{this.getFieldTitle(field)}</th>)}
                </tr>
                <tr className={`${styles.fields}`} id="fields">
                  {fields.map((field, fieldIndex) => <th className={this.setColumnHideClass("", fieldIndex)} key={fieldIndex}>{this.getFieldInput(field, filterField, filterValue)}</th>)}
                </tr>
              </thead>
              <tbody>
                {filteredItems.map((item, itemIndex) => <tr key={itemIndex}>{
                  fields.map((field, fieldIndex) => <td className={this.setColumnHideClass("", fieldIndex)} key={fieldIndex} onClick={() => this.onClickHandler(fieldIndex, item)}>{item[field.InternalName]}</td>)
                }</tr>)}
              </tbody>
              <tfoot>
                <tr><th colSpan={fields.length}><span className={`${styles.totop} pcursor`}>&#9650; TOP</span></th></tr>
              </tfoot>
            </table>
            <Slickmodal fields={this.state.fields} item={this.state.selectedItem} showModal={this.state.showModal} onClose={() => { this.setState({ showModal: false }) }} />
          </div>
        )}
      </section>
    );
  }

}
