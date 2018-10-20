import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  RowAccessor,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import EmailFileUrlComponent from '../../components/emailfileurlcomponent';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEmailFileUrlCommandSetProperties {
  
}

const LOG_SOURCE: string = 'EmailFileUrlCommandSet';

export default class EmailFileUrlCommandSet extends BaseListViewCommandSet<IEmailFileUrlCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized EmailFileUrlCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    
    const emailFileUrlCommand: Command = this.tryGetCommand('EmailFileUrl');

    if (emailFileUrlCommand) {
      Log.info(LOG_SOURCE, 'EmailFileUrl commandset not found');

      // This command should be hidden unless exactly one row is selected.
      if (event.selectedRows.length === 1) {
        emailFileUrlCommand.visible = true;
      }
      else
      {
        Log.info(LOG_SOURCE, 'More than one items are selected');
      }
    }
    else
    {
      Log.info(LOG_SOURCE, 'EmailFileUrl commandset not found');
    }
  }
  
  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.itemId) {
      case 'EmailFileUrl':
        this._showItemUrlDialog(event.selectedRows[0]);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  

  private _showItemUrlDialog(row: RowAccessor) {

    const div = document.createElement('div');

    const dialog: React.ReactElement<{}> = React.createElement(EmailFileUrlComponent, {
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listTitle: this.context.pageContext.list.title,
      itemId: row.getValueByName("ID"),
      fileName: row.getValueByName("FileName"),
      fileRelativePath: row.getValueByName("FileRef"),
      spHttpClient: this.context.spHttpClient
    });

    ReactDOM.render(dialog, div);

  }

}
