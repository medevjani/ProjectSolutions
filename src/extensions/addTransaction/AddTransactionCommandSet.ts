import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters
} from '@microsoft/sp-listview-extensibility';
import addTransaction from './components/dlgTaskTransaction';
import { spfi, SPFx } from "@pnp/sp";
import { IField, IFieldInfo } from "@pnp/sp/fields/types";
import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";
import * as appSettings from 'appSettings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAddTransactionCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'AddTransactionCommandSet';
let options: string[] = [];

export default class AddTransactionCommandSet extends BaseListViewCommandSet<IAddTransactionCommandSetProperties> {

  public async onInit(): Promise<void> {
    
    // Initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_AddTransaction');
    compareOneCommand.visible = false;

    await super.onInit();
    const sp = spfi().using(SPFx(this.context));
    
    // Get the choice values of status field
    const statusField: IFieldInfo = await sp.web.lists.getByTitle(appSettings.TaskTransactionListName).fields.getByInternalNameOrTitle("Status")();
    options = statusField.Choices;
    
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_AddTransaction':
        var userEmail: string = this.context.pageContext.user.email;
        
        // Initialize the component to a dialog box where user can add a new transaction
        const dialog: addTransaction = new addTransaction({ isBlocking: true, email: userEmail, contextConfig: this.context, selectedItem: event.selectedRows, data: options });
        dialog.show();

        break;
      default:
        throw new Error('Unknown command');
    }
  }
  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void { 
    debugger
    var Libraryurl = this.context.pageContext.list.title; 
      
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_AddTransaction'); 
    if (compareOneCommand) { 
      // This command should be hidden unless exactly one row is selected. 
      compareOneCommand.visible = (event.selectedRows.length === 1 && Libraryurl === appSettings.ProjectTaskListName ); 
    } 
  }
}
