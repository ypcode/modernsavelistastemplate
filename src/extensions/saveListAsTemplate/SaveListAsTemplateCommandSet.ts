import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { assign } from "office-ui-fabric-react";
import * as ReactDOM from "react-dom";
import * as React from "react";

import * as strings from 'SaveListAsTemplateCommandSetStrings';

import { SaveListAsSiteScriptPanel, ISaveListAsSiteScriptPanelProps } from "../components/SaveListAsSiteScriptPanel";
import { ContextServiceKey } from '../../services/ContextService';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISaveListAsTemplateCommandSetProperties {

}

const LOG_SOURCE: string = 'SaveListAsTemplateCommandSet';

export default class SaveListAsTemplateCommandSet extends BaseListViewCommandSet<ISaveListAsTemplateCommandSetProperties> {

  private _panelHostElement: HTMLElement;

  @override
  public onInit(): Promise<void> {
    return new Promise((resolve, reject) => {
      Log.info(LOG_SOURCE, 'Initialized SaveListAsTemplateCommandSet');
      this._panelHostElement = document.createElement("div");
      document.body.appendChild(this._panelHostElement);
      this.context.serviceScope.whenFinished(() => {
        const contextService = this.context.serviceScope.consume(ContextServiceKey);
        contextService.configure(this.context);
        resolve();
      });
    });
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const saveAsSiteScriptCommand: Command = this.tryGetCommand('SAVE_AS_SITE_SCRIPT');
    if (saveAsSiteScriptCommand) {
      // This command should be hidden when item selected and user is SP Admin.
      saveAsSiteScriptCommand.visible = event.selectedRows.length === 0;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'SAVE_AS_SITE_SCRIPT':
        this._showPanel();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _renderPanelComponent(props: any) {
    const element: React.ReactElement<ISaveListAsSiteScriptPanelProps> = React.createElement(SaveListAsSiteScriptPanel,
      assign(
        {
          onClose: () => this._hidePanel(),
          isOpen: false,
          listId: this.context.pageContext.list.id.toString(),
          listTitle: this.context.pageContext.list.title,
          serviceScope: this.context.serviceScope
        },
        props));

    ReactDOM.render(element, this._panelHostElement);
  }

  private _showPanel() {
    this._renderPanelComponent({ isOpen: true });
  }

  private _hidePanel() {
    this._renderPanelComponent({ isOpen: false });
  }
}
