import * as React from 'react';
import SidePanelExample from './SidePanelExample';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AvailablePanels } from '../../../Utils/Helpers';
import FolderSidePanelExample from './FolderSidePanelExample';
require('../../../styles/styles.css');

export interface ParentPanelProps {
  onClose: () => void;
  isOpen: boolean;
  context: WebPartContext;
  listItemId: number;
  panelVersion: string;
  docUrl: string;
  availablePanels: AvailablePanels;
}

export default class PanelParent extends React.Component<ParentPanelProps, any> {

  constructor(props: ParentPanelProps) {
    super(props);
  }

  public render(): React.ReactElement<ParentPanelProps> {

    return (
      <div id="mainPanel">
        {
          this.props.isOpen ? this.GetPanel() : null
        }
      </div>
    );
  }

  private GetPanel(): JSX.Element {
    let panelToDraw = null;
    switch (this.props.availablePanels) {
      case AvailablePanels.ReactionPanel:
        panelToDraw = <SidePanelExample onClose={this.props.onClose} context={this.props.context}
          version={this.props.panelVersion} docUrl={this.props.docUrl} />;
        break;

      case AvailablePanels.FolderPanel:
        panelToDraw = <FolderSidePanelExample onClose={this.props.onClose} context={this.props.context}
          version={this.props.panelVersion} docUrl={this.props.docUrl} listItemID={this.props.listItemId}/>;
        break;

      case AvailablePanels.Draft:
        panelToDraw = <SidePanelExample onClose={this.props.onClose} context={this.props.context}
          version={this.props.panelVersion} docUrl={this.props.docUrl} />;
        break;
    }

    return panelToDraw;

  }
}
