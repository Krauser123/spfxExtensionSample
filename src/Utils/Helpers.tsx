import { RowAccessor } from "@microsoft/sp-listview-extensibility";

export enum AvailablePanels {
  None = 0,
  ReactionPanel = 1,
  FolderPanel = 2,
  Draft = 3,
}

export enum ModerationStatus {
  Approved = 0,
  Rejected = 1,
  Draft = 3,
  Pending = 2,
  Scheduled = 4
}

export interface ReactionItem {
  Title: string;
  UserId: number;
  UserStringId: string;
  UserData: UserData;
  Reaction: string;
  DocUrl: string;
  Created: Date;
  Modified: Date;
  AuthorId: number;
  EditorId: number;
  Id: number;
}

export class UserData {
  public ID: number;
  public Email: string;
  public FullName: string;

  constructor(id: number, email: string, fullName: string) {
    this.ID = id;
    this.Email = email;
    this.FullName = fullName;
  }
}

export class ItemData {

  public Etag: string;
  public ListItemId: number;
  public ItemName: string;
  public FileURLEncoded: string;
  public DocId: string;

  constructor() {

  }

  public getItemPropertiesFromRowAccesor(selectedItem: RowAccessor) {
    if (selectedItem != null) {
      //Get item properties
      this.Etag = selectedItem.getValueByName('.etag');
      this.ListItemId = selectedItem.getValueByName('ID') as number;
      this.ItemName = selectedItem.getValueByName('FileLeafRef');
      this.FileURLEncoded = selectedItem.getValueByName("FileRef").replace(/ /g, "%20");

      //Document ID Service must be activated to use it
      this.DocId = selectedItem.getValueByName('_dlc_DocIdUrl.desc');
    }
  }
}
