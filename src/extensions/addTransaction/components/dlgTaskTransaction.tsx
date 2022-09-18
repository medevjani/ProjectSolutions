import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
    PrimaryButton,
    DialogFooter,
    DialogContent,
} from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
// Import Textfield component
import { TextField } from 'office-ui-fabric-react/lib/TextField';
// Import Button component
import { MessageBar, MessageBarType, IStackProps, Stack } from 'office-ui-fabric-react'
import styles from '../tasks.module.scss';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import * as appSettings from 'appSettings';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";
import "@pnp/sp/items";
const verticalStackProps: IStackProps = {  
    styles: { root: { overflow: 'hidden', width: '100%' } },  
    tokens: { childrenGap: 20 }  
  }; 
interface IDialogContentProps {
    message: string;
    close: () => void;
    data?: any[];
    email?: string;
    contextConfig?: any;
    selectedItem?: any;
    opType?: string;

}
const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 }
};

let options: IDropdownOption[] = [

];

class DialogLabelData extends React.Component<IDialogContentProps, any> {

    public data: any[];
    public email: string;
    public contextConfig: any;
    public selectedItem: any;
    public close: () => void;
    //Initialize all variables and bind functions
    constructor(props: IDialogContentProps | Readonly<IDialogContentProps>) {
        super(props);
        //debugger;
        this.data = props.data;
        this.email = props.email;
        this.contextConfig = props.contextConfig;
        this.selectedItem = props.selectedItem;
        this.close = props.close;
        var items: IDropdownOption[]=[];
        if(this.data && this.data.length > 0){
            this.data.forEach(element => {
                items.push({key:element, text:element});
            });
        }
        this.state = {           
            options: this.data && this.data.length > 0 ? items : [],
            selectedStatus: "",
            showSpinner: false,
            suppliedItems:0,
            selectedItem: this.selectedItem,
            hasSuppliedItemError:false,
            hasStatusError:false
        }
        this.addTaskTransaction = this.addTaskTransaction.bind(this);
        this.statusSelected = this.statusSelected.bind(this);
        this.onChangeSuppliedItems = this.onChangeSuppliedItems.bind(this);
        this.onSupplierItemsError = this.onSupplierItemsError.bind(this);
    }

    //Load dropdown with values set on initial load,in case its not loaded fetch data by calling the azure function again to get labels
    public componentDidMount() {
        //debugger;
        if (this.data && this.data.length > 0) {

        }
        else {
            
        }

    }
    // Modify state to set status from drop down
    public statusSelected(status: IDropdownOption) {
        //debugger;
        this.setState({
            selectedStatus: status.text
        })
    }

    public onChangeSuppliedItems(event: any) {
        //debugger;
        this.setState({
            suppliedItems: event.target.value
        })
        
    }
    public onSupplierItemsError(suppliedItemsData: any) {
        debugger;
        if(suppliedItemsData === ""){
            this.setState({
                hasSuppliedItemError: true
            })
            return "Please enter suppplied item count!!";
        }
        else{
            this.setState({
                hasSuppliedItemError: false
            })
            return ""
        }
    }
    
    // Add task transaction to transaction list and update total supplied items in parent list
    public async addTaskTransaction() {
        //debugger;
        if(this.state.suppliedItems==="" || this.state.selectedStatus===""){
            this.setState({  
                message: "Please fill up the details before submitting the details.",  
                showMessageBar: true,  
                messageType: MessageBarType.error  
              });  
        }
        else{
        const sp = spfi().using(SPFx(this.contextConfig));
        try {  
            // Add item to the transaction list
            await sp.web.lists.getByTitle(appSettings.TaskTransactionListName).items.add({  
              Title: this.state.title,
              ProjectId: this.state.selectedItem[0]._values.get("Project")[0].lookupId,
              ItemId: this.state.selectedItem[0]._values.get("Item")[0].lookupId,
              ItemCategoryId: this.state.selectedItem[0]._values.get("ItemCategory")[0].lookupId,
              SuppliedItems:this.state.suppliedItems,
              Status:this.state.selectedStatus
            });
            // Update item of project task list
            await sp.web.lists.getByTitle(appSettings.ProjectTaskListName).items.getById(this.state.selectedItem[0]._values.get("ID")).update({  
                TotalSuppliedItems:parseInt(this.state.selectedItem[0]._values.get("TotalSuppliedItems"),10) + parseInt(this.state.suppliedItems,10)
              }).then(_ => {  
                        // Reload the page to show the updated value of TotalSuppliedItems field in the view
                        location.reload();
                    });
            this.setState({  
              message: "Task transaction is added successfully! Total supplied items value is also updated.",  
              showMessageBar: true,  
              messageType: MessageBarType.success  
            });  
          }  
          catch (error) {  
            this.setState({  
              message: "Item  creation/updation failed with error: " + error,  
              showMessageBar: true,  
              messageType: MessageBarType.error  
            });  
          }
        }
    }

    // Page event that renders the component with the form fields
    public render(): JSX.Element {
        //debugger;
        let email = this.email;
        return <DialogContent
            title=''
            subText={this.props.message}
            onDismiss={this.props.close}
            showCloseButton={true}
        >
            {  
                this.state.showMessageBar  
                ?  
                <div className="form-group">  
                    <Stack {...verticalStackProps}>  
                    <MessageBar messageBarType={this.state.messageType}>{this.state.message}</MessageBar>  
                    </Stack>  
                </div>  
                :  
                null  
            }  
            <div className={styles.header}>Project:</div>
            <div className={styles.details}>{this.selectedItem[0]._values.get("Project")[0].lookupValue}</div>

            <div className={styles.header}>Project Category:</div>
            <div className={styles.details}>{this.selectedItem[0]._values.get("ItemCategory")[0].lookupValue}</div>

            <div className={styles.header}>Item:</div>
            <div className={styles.details}>{this.selectedItem[0]._values.get("Item")[0].lookupValue}</div>

            <div className={styles.header}>Supllied Items:</div>
            <TextField type="number"
                  value={this.state.suppliedItems}
                  onChange={this.onChangeSuppliedItems}
                  required={true}
                  onGetErrorMessage={this.onSupplierItemsError}
                  width={50}
            />

            <div className={styles.header}>Status:</div>
            <Dropdown placeholder="Staus" 
                options={this.state.options} 
                styles={dropdownStyles} 
                onChanged={this.statusSelected} 
                dropdownWidth={375} 
                required={true}
            />

            <DialogFooter>
                <PrimaryButton text='Add Transaction' title='Add Transaction' onClick={this.addTaskTransaction} />
            </DialogFooter>
        </DialogContent>;
    }
}

// Class which creates the dialog container to add transaction
export default class addTransaction extends BaseDialog {
    public message: string;
    public data: any[];
    public email: string;
    public contextConfig: any;
    public selectedItem: any;
    public opType: string;
    constructor(config: any) {
        super(config);
        //debugger;
        this._close = this._close.bind(this);
        this.email = config.email;
        this.contextConfig = config.contextConfig;
        this.selectedItem = config.selectedItem;
        this.opType = config.opType;
        this.data = config.data;
    }

    //Closes dialog object for apply label
    private _close() {
        this.close();
        ReactDOM.unmountComponentAtNode(this.domElement);
    }

    // Sets all dialog values with label properties
    public render(): void {
        ReactDOM.render(<DialogLabelData
            close={this.close}
            message={this.message}
            data={this.data}
            email={this.email}
            contextConfig={this.contextConfig}
            selectedItem={this.selectedItem}
        />, this.domElement);
    }

}