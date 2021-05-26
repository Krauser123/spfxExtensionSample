import * as React from 'react';
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
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode, SelectChildrenMode } from "@pnp/spfx-controls-react/lib/TreeView";
import { IIconProps, PrimaryButton, DialogFooter, Panel, PanelType } from "office-ui-fabric-react";
import { Items } from '@pnp/sp/items';

export interface IFolderSidePanelExampleState {
  FileUrl: string;
  Disabled?: boolean;
  ListItemID: number;
  ShowDataPanel: boolean;
}

export interface IFolderSidePanelExampleProps {
  onClose: () => void;
  isOpen?: boolean;
  context: WebPartContext;
  listItemID: number;
  version: string;
  docUrl: string;
}

export default class FolderSidePanelExample extends React.Component<IFolderSidePanelExampleProps, IFolderSidePanelExampleState> {
  private isVersionDrawn: boolean = false;
  private UserReactionList_Name: string = "UserReactionList";
  private skypeCheckIcon: IIconProps = { iconName: 'SkypeCheck' };

  constructor(props: IFolderSidePanelExampleProps) {

    super(props);

    //Binding
    this._onCancel = this._onCancel.bind(this);

    this.state = {
      FileUrl: null,
      Disabled: false,
      ListItemID: props.listItemID,
      ShowDataPanel: false,
    };

    //Load data
    this._onRead().then();
  }

  private _onCancel() {
    this.props.onClose();
  }

  private async _onRead() {
    //Load Data from SPList
    await this.getReactionData();
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
      ShowDataPanel: true
    });
  }

  public render(): React.ReactElement<IFolderSidePanelExampleProps> {

    return (
      <Panel id="sidePanel" isOpen={true} type={PanelType.medium} isFooterAtBottom={true} onRenderFooterContent={this._onRenderFooterContent} >
        {Utils.GetLoadingRoller(this.state.ShowDataPanel)}

        <div id="dataPanel" style={this.state.ShowDataPanel == true ? { display: 'block' } : { display: 'none' }}>
          <div style={{ textAlign: "center" }}>
            <h2>{strings.FolderPanelTitle}</h2>
          </div>
          <div id="treeView">
            <TreeView
              items={this.getTreeViewItem()}
              defaultExpanded={false}
              selectionMode={TreeViewSelectionMode.Multiple}
              selectChildrenMode={SelectChildrenMode.Select | SelectChildrenMode.Unselect}
              showCheckboxes={true}
              treeItemActionsDisplayMode={TreeItemActionsDisplayMode.ContextualMenu}
              defaultSelectedKeys={['key1', 'key2']}
              expandToSelected={true}
              defaultExpandedChildren={true}
              onSelect={this.onTreeItemSelect}
              onExpandCollapse={this.onTreeItemExpandCollapse}
            />
          </div>
        </div>
      </Panel>
    );
  }

  private _onRenderFooterContent = (): JSX.Element => {

    try { //Assign onClose to upper [X]  button
      let upperCloseButton = document.getElementsByClassName("ms-PanelAction-close")[0] as HTMLElement;
      upperCloseButton.onclick = () => { this._onCancel(); };

      //Draw Version
      this.isVersionDrawn = Utils.drawVersion(this.isVersionDrawn, this.props.version);
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

  private onTreeItemSelect(items: ITreeItem[]) {
    console.log("Items selected: ", items);
  }

  private onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item);
  }

  private getLinks = async () => {
    // gets list's folders
    const listFolders = await sp.web.lists.getByTitle("My List").rootFolder.folders();

    // gets item's folders
    const itemFolders = await sp.web.lists.getByTitle("My List").items.getById(1).folder.folders();

  }

  private getTreeViewItem = (): any => {
    let items = [
      {
        key: "R1",
        label: "Root",
        subLabel: "This is a sub label for node",
        iconProps: this.skypeCheckIcon,
        actions: [{
          title: "Get item",
          iconProps: {
            iconName: 'Warning',
            style: {
              color: 'salmon',
            },
          },
          id: "GetItem",
          actionCallback: async (treeItem: ITreeItem) => {
            console.log(treeItem);
          }
        }],
        children: [
          {
            key: "1",
            label: "Parent 1",
            selectable: false,
            children: [
              {
                key: "3",
                label: "Child 1",
                subLabel: "This is a sub label for node",
                actions: [{
                  title: "Share",
                  iconProps: {
                    iconName: 'Share'
                  },
                  id: "GetItem",
                  actionCallback: async (treeItem: ITreeItem) => {
                    console.log(treeItem);
                  }
                }],
                children: [
                  {
                    key: "gc1",
                    label: "Grand Child 1",
                    actions: [{
                      title: "Get Grand Child item",
                      iconProps: {
                        iconName: 'Mail'
                      },
                      id: "GetItem",
                      actionCallback: async (treeItem: ITreeItem) => {
                        console.log(treeItem);
                      }
                    }]
                  }
                ]
              },
              {
                key: "4",
                label: "Child 2",
                iconProps: this.skypeCheckIcon
              }
            ]
          },
          {
            key: "2",
            label: "Parent 2"
          },
          {
            key: "5",
            label: "Parent 3",
            disabled: true
          },
          {
            key: "6",
            label: "Parent 4",
            selectable: true
          }
        ]
      },
      {
        key: "R2",
        label: "Root 2",
        children: [
          {
            key: "8",
            label: "Parent 5"
          }
        ]
      }
    ];

    return items;
  }
}
