declare interface IExtensionSampleCommandSetStrings {  
  Accept: string;
  Cancel: string;
  Close: string;
  ReactionComments: string;
  ReactionPanelTitle: string;
  FolderPanelTitle: string;
  NoItems: string;
}

declare module 'ExtensionSampleCommandSetStrings' {
  const strings: IExtensionSampleCommandSetStrings;
  export = strings;
}
