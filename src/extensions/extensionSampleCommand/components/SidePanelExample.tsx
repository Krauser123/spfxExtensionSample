import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users";
import * as strings from 'ExtensionSampleCommandSetStrings';
import { ReactionItem } from '../../../Utils/Helpers';
import { Utils } from '../../../Utils/Utilities';
import {
  TextField, PrimaryButton, IContextualMenuProps, DefaultButton, DialogFooter, Panel, PanelType, ActivityItem, Link
} from "office-ui-fabric-react";

export interface ISidePanelExampleState {
  FileUrl: string;
  Disabled?: boolean;
  ShowDataPanel: boolean;
  ReactionItems: ReactionItem[];
  OptionsForButton: IContextualMenuProps;
  ReactionComments: string;
  SelectedOptionForButton: string;
}

export interface ISidePanelExampleProps {
  onClose: () => void;
  isOpen?: boolean;
  context: WebPartContext;
  version: string;
  docUrl: string;
}

export default class SidePanelExample extends React.Component<ISidePanelExampleProps, ISidePanelExampleState> {
  private isVersionDrawn: boolean = false;
  private UserReactionList_Name: string = "UserReactionList";

  constructor(props: ISidePanelExampleProps) {

    super(props);

    //Binding
    this._onCancel = this._onCancel.bind(this);

    this.state = {
      FileUrl: null,
      Disabled: false,
      ShowDataPanel: false,
      ReactionItems: [],
      OptionsForButton: this.buildButtonOptions(null),
      ReactionComments: "",
      SelectedOptionForButton: "Select a reaction"
    };

    //Load data
    this._onRead().then();
  }

  private buildButtonOptions = (items: any) => {

    let menuProps: IContextualMenuProps = { items: [] };

    if (items != null) {

      for (let i = 0; i < items.Choices.length; i++) {
        let choiceOption = items.Choices[i];
        menuProps.items.push({
          key: choiceOption,
          text: choiceOption,
          onClick: this.splitButtonClick
        }
        );
      }
    }
    return menuProps;
  }

  private splitButtonClick = (event) => {
    this.setState({ SelectedOptionForButton: event.target.innerText });
  }

  private _onReactionCommentsChanged = (reactionComments: any) => {
    this.setState({ ReactionComments: reactionComments.target.value });
  }

  private _onCancel() {
    this.props.onClose();
  }

  private async _onRead() {

    //Load Data from SPList
    await this.getReactionData();

    await this.getChoicesValues();
  }

  private getReactionData = async () => {
    let reactionItems: ReactionItem[];
    try {
      let itemInSPList: ReactionItem[] = await sp.web.lists.getByTitle(this.UserReactionList_Name).items.filter("DocURL eq '" + this.props.docUrl + "'").getAll();

      for (let i = 0; i < itemInSPList.length; i++) {
        itemInSPList[i].UserData = await Utils.getUserDataFromUserID(this.props.context, itemInSPList[i].UserId);
      }

      reactionItems = itemInSPList;
    }
    catch (ex) {
      console.error(ex);
    }

    //Update state
    this.setState({
      ShowDataPanel: true,
      ReactionItems: reactionItems
    });
  }

  public render(): React.ReactElement<ISidePanelExampleProps> {

    let activityItems = [];
    if (this.state.ReactionItems.length > 0) {
      activityItems = this.createActivityItemsFromSPData(this.state.ReactionItems);
    } else {
      this.getNoItemsPanel();
    }

    return (
      <Panel id="sidePanel" isOpen={true} type={PanelType.medium} isFooterAtBottom={true} onRenderFooterContent={this._onRenderFooterContent} >
        {Utils.GetLoadingRoller(this.state.ShowDataPanel)}

        <div id="dataPanel" style={this.state.ShowDataPanel == true ? { display: 'block' } : { display: 'none' }}>
          <div style={{ textAlign: "center" }}>
            <h2>{strings.ReactionPanelTitle}</h2>
          </div>
          <div id="activityDiv">
            {activityItems.map((item: { key: string | number }) => (
              <ActivityItem {...item} key={item.key} className="mainActivity" />
            ))}
          </div>

          <div className="reactionForm">
            <TextField placeholder={strings.ReactionComments} multiline rows={4} value={this.state.ReactionComments} onChange={this._onReactionCommentsChanged} />

            <div className="buttonOptionsContainer">
              <DefaultButton
                className="buttonOptions"
                text={this.state.SelectedOptionForButton}
                split
                splitButtonAriaLabel="See options"
                aria-roledescription="split button"
                menuProps={this.state.OptionsForButton}
                onClick={this.addReactionToSPList}
                disabled={false}
                checked={true}
                onChange={this.OnChageButton}
              />
            </div>
          </div>
        </div>
      </Panel>
    );
  }

  private getChoicesValues = async () => {

    let itemInSPList = await sp.web.lists.getByTitle(this.UserReactionList_Name).fields.getByInternalNameOrTitle('Reaction').select('Choices,ID').get();
    let options = this.buildButtonOptions(itemInSPList);
    this.setState({ OptionsForButton: options });
  }

  private _onRenderFooterContent = (): JSX.Element => {

    try { //Assign onClose to upper [X]  button
      let upperCloseButton = document.getElementsByClassName("ms-PanelAction-close")[0] as HTMLElement;
      upperCloseButton.onclick = () => { this._onCancel(); };

      //Draw Version
      this.isVersionDrawn=Utils.drawVersion(this.isVersionDrawn, this.props.version);
    }
    catch (err) {
      //This set can fails in some renders, IE: panel not visible yet
    }

    return (
      <DialogFooter>
        <PrimaryButton text={strings.Close} onClick={this._onCancel} checked={true} />
      </DialogFooter>
    );
  }

  private createActivityItemsFromSPData = (reactionItems: ReactionItem[]): any[] => {
    let activityItems = [];

    for (let i = 0; i < reactionItems.length; i++) {
      let reactionItem = reactionItems[i];

      //Add item to collection
      activityItems.push(this.createActivityItem(i, reactionItem));
    }

    return activityItems;
  }

  private createActivityItem = (index: number, itemToDraw: any): any => {
    return {
      key: index,
      activityDescription: [
        <Link
          key={index}
          className="userName"
          onClick={() => { window.open("/_layouts/15/me.aspx/?p=" + itemToDraw.UserData.Email); }}>
          {itemToDraw.UserData.FullName}
        </Link>,
        <span key={index + 1}> {itemToDraw.Reaction} </span>,
      ],
      activityPersonas: [{ imageUrl: this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?AccountName=" + itemToDraw.UserData.Email + "&size=s" }],
      comments: itemToDraw.Title,
      timeStamp: Utils.adjustTimeZone(itemToDraw.Created.toString())
    };
  }

  private OnChageButton = (event: any) => {
    console.info(event);
  }

  private getNoItemsPanel = () => {
    try {
      let rootElement = document.getElementById("mainPanel");

      let noItemsPanel = (
        <div id="mainPanel">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md6" style={{ "textAlign": "center" }}>
              <label style={{ fontWeight: 600 }}>{strings.NoItems}</label>
            </div>
          </div>
        </div>);

      ReactDOM.render(noItemsPanel, rootElement);
    }
    catch (ex) {
      console.error(ex);
    }
  }

  private addReactionToSPList = async () => {

    try {
      await sp.web.lists.getByTitle(this.UserReactionList_Name).items.add({
        Title: this.state.ReactionComments,
        UserId: (await sp.web.currentUser()).Id,
        Reaction: this.state.SelectedOptionForButton,
        DocURL: this.props.docUrl
      });
    }
    catch (ex) {
      console.error(ex);
    }
  }
}
