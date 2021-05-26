import { override } from '@microsoft/decorators';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { assign } from '@uifabric/utilities';
import { sp } from "@pnp/sp";
import {
  BaseListViewCommandSet, Command,
  IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters, RowAccessor
} from '@microsoft/sp-listview-extensibility';
import ParentPanel, { ParentPanelProps } from './components/ParentPanel';
import { AvailablePanels, ItemData, ModerationStatus } from '../../Utils/Helpers';

export interface IExtensionSampleCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
  listItemId: number;
}

const LOG_SOURCE: string = 'ExtensionSampleCommandSet';
const VERSION: string = "1.0.1";

const COMMAND_BTN_1 = "COMMAND_BTN_1";
const COMMAND_BTN_3 = "COMMAND_BTN_3";
const COMMAND_BTN_4 = "COMMAND_BTN_4";


export default class ExtensionSampleCommandSet extends BaseListViewCommandSet<IExtensionSampleCommandSetProperties> {

  private panelDomElement: HTMLDivElement;
  private commandOne: Command = null;
  private commandThree: Command = null;
  private commandFour: Command = null;

  @override
  public onInit(): Promise<void> {
    this._dismissPanel = this._dismissPanel.bind(this);

    console.info(LOG_SOURCE, 'Initialized v.' + VERSION);

    // Setup the PnP JS with SPFx context
    sp.setup({
      spfxContext: this.context
    });

    this.panelDomElement = document.body.appendChild(document.createElement("div"));

    //Get commands
    this.commandOne = this.tryGetCommand(COMMAND_BTN_1);
    this.commandThree = this.tryGetCommand(COMMAND_BTN_3);
    this.commandFour = this.tryGetCommand(COMMAND_BTN_4);

    //Check visibilities
    this.checkCommandSetVisibilities(null);

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    this.checkCommandSetVisibilities(event);
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    let panelToDraw: AvailablePanels;

    //We need to know in that button
    switch (event.itemId) {
      case COMMAND_BTN_1:
        panelToDraw = AvailablePanels.ReactionPanel;
        break;

      case COMMAND_BTN_3:
        panelToDraw = AvailablePanels.FolderPanel;
        break;

      case COMMAND_BTN_4:
        panelToDraw = AvailablePanels.Draft;
        break;
    }

    //Only one task at each time
    let selectedItem = event.selectedRows[0];

    //Show Panel
    this._renderPanelComponent(selectedItem, panelToDraw);
  }

  private _dismissPanel() {
    this._renderPanelComponent(null, AvailablePanels.None);
  }

  private _renderPanelComponent(selectedItem: RowAccessor, availablePanels: AvailablePanels) {
    let itemProperties: ItemData = new ItemData();
    itemProperties.getItemPropertiesFromRowAccesor(selectedItem);

    let element: React.ReactElement<ParentPanelProps> = React.createElement(
      ParentPanel, assign({
        onClose: this._dismissPanel,
        currentTitle: null,
        listItemId: itemProperties.ListItemId,
        isOpen: selectedItem != null ? true : false,
        listId: this.context.pageContext.list.id,
        default: this.context.listView,
        docUrl: itemProperties.FileURLEncoded,
        context: this.context,
        panelVersion: VERSION,
        availablePanels: availablePanels
      }, {}));

    //Draw on DOM
    ReactDom.render(element, this.panelDomElement);
  }

  private checkIfCommandMustBeVisible(commandButon: any, event: IListViewCommandSetListViewUpdatedParameters, numberOfSelectedRows: number, contentTypeAvailable: string[]): boolean {
    let isVisible: boolean;

    let contentType = event.selectedRows[0].getValueByName('ContentType');
    if (commandButon && event.selectedRows.length === numberOfSelectedRows && contentTypeAvailable.indexOf(contentType) > -1) {
      isVisible = true;
    }

    return isVisible;
  }

  private checkCommandSetVisibilities = (event: IListViewCommandSetListViewUpdatedParameters) => {
    try {
      if (event != null && event.selectedRows != null && event.selectedRows.length > 0) {
        // This command should be hidden unless that one row is selected and type is Doc
        this.commandOne.visible = this.checkIfCommandMustBeVisible(this.commandOne, event, 1, ["Document"]);

        // This command should be hidden unless that one row is selected and type is Folder
        this.commandThree.visible = this.checkIfCommandMustBeVisible(this.commandThree, event, 1, ["Folder"]);

        // This command should be hidden unless that one row is selected and type is Folder
        this.commandFour.visible = this.checkModerationStatusForFiles(event.selectedRows[0], ModerationStatus.Approved);
      }
      else {
        this.commandOne.visible = this.commandThree.visible = this.commandFour.visible = false;
      }
    }
    catch (ex) {
      console.error(ex);
    }
  }

  private checkModerationStatusForFiles(selectedRow: any, moderationStatus: ModerationStatus) {
    let isVisible: boolean;
    let fileType: string = selectedRow.getValueByName('File_x0020_Type');
    let docModerationStatus: number = parseInt(selectedRow.getValueByName('_ModerationStatus.'));

    if (fileType != "" && docModerationStatus == moderationStatus) {
      isVisible = true;
    }

    return isVisible;
  }
}
