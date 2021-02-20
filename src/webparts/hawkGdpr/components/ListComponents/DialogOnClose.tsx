import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton} from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, DayOfWeek, IDatePickerStrings, mergeStyleSets, IColumn } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp-commonjs";
import { PivotTabsLargeExample } from '../PivotLargeTabsExample';

let dialogClosed = false;
export const DayPickerStrings: IDatePickerStrings = {
    months: [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December',
    ],
  
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  
    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker',
   // monthPickerHeaderAriaLabel: '{0}, select to change the year',
   // yearPickerHeaderAriaLabel: '{0}, select to change the month',
  };
  
  const controlClass = mergeStyleSets({
    control: {
      margin: '0 0 15px 0',
      maxWidth: '300px',
    },
  });
  
const dialogContentProps = {
    type: DialogType.normal,
    //title: 'Comments',
    //subText: 'Plese provide the comments in the TextBox and Click on Save. Or Else Click Cancel to close the Dialog',
  };
  
  const modelProps = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
  };

export interface IDialogOnCloseProps{
    hideDialog : boolean;
    savedCaseData : any;
    modifyState : Function;
    allItems :any[];
    currentUserName : string;
    caseType : string;
    retainedItems? : any[];
    normalItems? : any[];
    migratedItems? : any[];
    allRetainedItems? : any[];
    retainedCaseColumns? : IColumn[];

}

export interface IDialogOnCloseState{
    ishidden : boolean;
    inputValue : string;
    selectedClosedDate : Date;
    retainJustification : string;
    valid : boolean;
    selectedNextReviewDate : Date;
}

export class DialogOnClose extends React.Component<IDialogOnCloseProps, IDialogOnCloseState>{
    
   private static _changed : boolean;
   private _validClosedDate : boolean;
   private _validURN : boolean;
   private _validRetainJustification : boolean;
   private _validNextReviewDate : boolean;
   private _allItems : any[];
   
   
    constructor(props: IDialogOnCloseProps){
        super(props);
        //this._allItems = this.props.allItems;
        console.log("dialog closed::"+DialogOnClose._changed);
        this._validClosedDate = true;
        this._validURN = true;
        this._validRetainJustification = true;
        this._validNextReviewDate = true;
        this.state = {
            ishidden : this.props.hideDialog,
            inputValue : "",
            selectedClosedDate : null,
            retainJustification : "",
            valid : true,
            selectedNextReviewDate : null
        };
    }

    private _close = () : void => {
        console.log("inside _close..");
        dialogClosed = true;
        DialogOnClose._changed = false;
        this._validClosedDate = true;
        this._validURN = true;
        this._validRetainJustification = true;
        this._validNextReviewDate = true;
        this.setState({
            ishidden : true,
            inputValue :'',
            selectedClosedDate : null,
            selectedNextReviewDate : null,
            retainJustification : ''
        });
        
        
    }

    private _validate = () : void => {
      
      if(this.props.savedCaseData.isChecked === true){
        if(!this.state.selectedClosedDate || this.state.selectedClosedDate === null){
          console.log("Closed Date is not provided..");
          this._validClosedDate = false;
        }else{
          console.log("Closed Date is provided..");
          this._validClosedDate = true;
        }
      }
      if(this.props.savedCaseData.isDeleteSelected === true){
        if(!this.state.inputValue || this.state.inputValue === null || this.state.inputValue === ''){
          console.log("URN is not provided..");
          this._validURN = false;
          
        }else{
          console.log("URN is provided..");
          this._validURN = true;
        }
      }

      if(this.props.savedCaseData.isRetainSelected === true){
        if(!this.state.retainJustification || this.state.retainJustification === null || 
          this.state.retainJustification === ''){
          console.log("Retain Justification is not provided..");
          this._validRetainJustification = false;
        }else{
          console.log("Retain Justification is provided..");
          this._validRetainJustification = true;
        }
        if(this.props.caseType === 'Retained' && (!this.state.selectedNextReviewDate ||
                                      this.state.selectedNextReviewDate === null)){
          console.log("Selected Next Review Date is not provided..");
          this._validNextReviewDate = false;
        }else{
          console.log("Selected Next Review Date is provided..");
          this._validNextReviewDate = true;
        }
      }
      
    }
    private _save = () : void => {
        console.log("inside _Save..");
        //ev.preventDefault();
        this._validate();
        
        console.log("this._validClosedDate ::"+this._validClosedDate);
        console.log("this._validURN::"+this._validURN);
        console.log("this._validRetainJustification::"+this._validRetainJustification);
        console.log("this._validNextReviewDate::"+this._validNextReviewDate);
        if(this._validClosedDate === true && this._validURN === true && 
          this._validRetainJustification === true && this._validNextReviewDate === true){

        
        dialogClosed = true;
        //variables for Close Case
        let status;
        let closedBy;
        let closedDate: Date = null;
        //variable for Flagging case for Delete or Retain
        let retainedOrDeleted;

        //variables for Deletion
        let urn;
        let deleteConsentDate;
        let deleteConsentBy;

        //variables for Retain
        let retainJustification;
        let retainConsentDate;
        let retainConsentBy;
        let nextReviewDate;

        

        DialogOnClose._changed = false;
        console.log(this.props.savedCaseData);
        let caseId = this.props.savedCaseData.caseId;
        console.log("value in save::"+this.state.inputValue);
        console.log("value of date in save ::"+this.state.selectedClosedDate);
        //close Case scenario
        if(this.props.savedCaseData.isChecked === true){
          status = 'Closed';
          closedBy = this.props.currentUserName;
          closedDate = this.state.selectedClosedDate;
        }else{
          status = this.props.savedCaseData.status;
          closedBy = this.props.savedCaseData.closedBy;
          if(this.props.savedCaseData.closedDate === ""){
            closedDate = null;
          }else{
            closedDate = this.props.savedCaseData.closedDate;
          }
          
        }
       
        if(this.props.savedCaseData.isDeleteSelected === true || 
            this.props.savedCaseData.isRetainSelected === true){
              //Flag Case for Deletion scenario
              if(this.props.savedCaseData.isDeleteSelected === true){
                retainedOrDeleted ='Delete';
                urn = this.state.inputValue;
                deleteConsentDate = new Date();
                deleteConsentBy = this.props.currentUserName;
              }
              //Flag Case for Retain scenario
              else if(this.props.savedCaseData.isRetainSelected === true){
                retainedOrDeleted = 'Retain';
                retainJustification = this.state.retainJustification;
                retainConsentDate = new Date();
                retainConsentBy = this.props.currentUserName;
                nextReviewDate = this.state.selectedNextReviewDate;
              }
          
        }else{
          retainedOrDeleted = this.props.savedCaseData.retainedOrDeleted;
          //delete scenario
          urn = this.props.savedCaseData.urn;
          if(this.props.savedCaseData.deleteConsentDate === ""){
            deleteConsentDate = null;
          }else{
            deleteConsentDate = this.props.savedCaseData.deleteConsentDate;
          }
          
          deleteConsentBy = this.props.savedCaseData.deleteConsentBy;

          //retain scenario
          retainJustification = this.props.savedCaseData.retainJustification;
          nextReviewDate = this.props.savedCaseData.nextReviewDate;
          if(this.props.savedCaseData.retainConsentDate === ""){
            retainConsentDate = null;
          }else{
            retainConsentDate = this.props.savedCaseData.retainConsentDate;
          }
          retainConsentBy = this.props.savedCaseData.retainConsentBy;
        }

        
        
        console.log("**********Close case*************");
        console.log("status in save ::"+status);
        console.log("closed By in save ::"+closedBy);
        console.log("closed date in save ::"+closedDate);
        console.log("**************Delete scenario*****");
        console.log("retained or deleted ::"+retainedOrDeleted);
        console.log("URN ::"+urn);
        console.log("delete consent date ::"+deleteConsentDate);
        console.log("delete consent By::"+deleteConsentBy);
        console.log("**********Retain Scenario***********");
        console.log("Retain Justification::"+retainJustification);
        console.log("Retain Consent Date ::"+retainConsentDate);
        console.log("Retain Consent By::"+retainConsentBy);
        let singleCaseItem: any[] = [];
         
          const list = sp.web.lists.getByTitle("Case");
         list.items.getById(caseId).update({
           Status : status,
           ClosedDate: closedDate,
           ClosedBy: closedBy,
           Retained_Or_Deleted : retainedOrDeleted,
           URN : urn,
           DeleteConsentDate : deleteConsentDate,
           DeleteConsentBy : deleteConsentBy,
           RetainJustification : retainJustification,
           RetainConsentDate : retainConsentDate,
           RetainConsentBy : retainConsentBy,
           NextReviewDate : nextReviewDate
         }).then( updatedItem => {
          list.items.getById(caseId).get()
          .then(caseItem => {
                  singleCaseItem.push(caseItem);
                  console.log("***updated item via promise....and length is ::"+singleCaseItem.length);
                  console.log(singleCaseItem); 
          }).then( () => {
            //For Normal TAB scenario, remove Retained case from Normal Case List and push it to Retained Case List
            if(this.props.caseType === 'Normal'){
              let normalCaseItems : any[] = [];
              let retainedCaseItems : any[] = [];
              let allRetainedItems : any[] = [];
              let retainedCaseColumns : IColumn[] = [];
              //Get existing Retained Case Items, if there are any
              console.log("props retained ITEMS in Normal Case...");
              console.log(this.props.retainedItems);
              if(this.props.retainedItems && this.props.retainedItems.length > 0){
                retainedCaseItems = this.props.retainedItems;
              }
              if(this.props.allRetainedItems && this.props.allRetainedItems.length > 0){
                allRetainedItems = this.props.allRetainedItems;
              }
              this._allItems.forEach((caseItem : any) => {
                if(caseItem.caseId === caseId){
                 console.log("caseId ::"+caseItem.caseId+" status is ::"+singleCaseItem[0].Status);
                 console.log("closed date::"+singleCaseItem[0].ClosedDate+" closed By::"+singleCaseItem[0].ClosedBy);
                  caseItem.status = singleCaseItem[0].Status;
                  caseItem.retainedOrDeleted = singleCaseItem[0].Retained_Or_Deleted;
                  caseItem.closedDate = singleCaseItem[0].ClosedDate;
                  caseItem.closedBy = singleCaseItem[0].ClosedBy;
                  caseItem.urn = singleCaseItem[0].URN;
                  caseItem.deleteConsentDate = singleCaseItem[0].DeleteConsentDate;
                  caseItem.deleteConsentBy = singleCaseItem[0].DeleteConsentBy;
                  if(caseItem.retainedOrDeleted === 'Retain'){
                    caseItem.retainJustification = singleCaseItem[0].RetainJustification;
                    caseItem.retainConsentDate = singleCaseItem[0].RetainConsentDate;
                    caseItem.retainConsentBy = singleCaseItem[0].RetainConsentBy;
                    retainedCaseItems.push(caseItem);
                    allRetainedItems.push(caseItem);
                  }else{
                    normalCaseItems.push(caseItem);
                  }
                }else{
                  normalCaseItems.push(caseItem);
                }
    
              });         
              console.log("after modification normal Items::");
              console.log(normalCaseItems);
              console.log("after modification retained Items==");
              retainedCaseColumns = this.props.retainedCaseColumns.slice();
              retainedCaseColumns.forEach(column => {
                if(column.fieldName === 'dueDate'){
                  column.isSorted = true;
                  column.isSortedDescending = true;
                }else{
                  column.isSorted = false;
                  column.isSortedDescending = false;
                }
              });
              retainedCaseItems = PivotTabsLargeExample._copyAndSort(retainedCaseItems, 'dueDate',true);
              console.log(retainedCaseItems);
              console.log("after modification allRetained items are ..");
              console.log(allRetainedItems);
              this.props.modifyState(normalCaseItems, allRetainedItems,retainedCaseItems, true, retainedCaseColumns);
            }
            //For Migrated TAB Scenario
            else if(this.props.caseType === 'Migrated'){
              let migratedCaseItems : any[] = [];
              let retainedCaseItems : any[] = [];
              let allRetainedItems : any[] = [];
              let retainedCaseColumns : IColumn[] = [];
              //Get existing Retained Case Items, if there are any
              console.log("props retained ITEMS in Migrated Case...");
              console.log(this.props.retainedItems);
              if(this.props.retainedItems && this.props.retainedItems.length > 0){
                retainedCaseItems = this.props.retainedItems;
              }
              if(this.props.allRetainedItems && this.props.allRetainedItems.length > 0){
                allRetainedItems = this.props.allRetainedItems;
              }
              this._allItems.forEach((caseItem : any) => {
                if(caseItem.caseId === caseId){
                 console.log("migrated caseId ::"+caseItem.caseId+" status is ::"+singleCaseItem[0].Status);
                 console.log("closed date::"+singleCaseItem[0].ClosedDate+" closed By::"+singleCaseItem[0].ClosedBy);
                  caseItem.status = singleCaseItem[0].Status;
                  caseItem.retainedOrDeleted = singleCaseItem[0].Retained_Or_Deleted;
                  caseItem.closedDate = singleCaseItem[0].ClosedDate;
                  caseItem.closedBy = singleCaseItem[0].ClosedBy;
                  caseItem.urn = singleCaseItem[0].URN;
                  caseItem.deleteConsentDate = singleCaseItem[0].DeleteConsentDate;
                  caseItem.deleteConsentBy = singleCaseItem[0].DeleteConsentBy;
                  if(caseItem.retainedOrDeleted === 'Retain'){
                    caseItem.retainJustification = singleCaseItem[0].RetainJustification;
                    caseItem.retainConsentDate = singleCaseItem[0].RetainConsentDate;
                    caseItem.retainConsentBy = singleCaseItem[0].RetainConsentBy;
                    retainedCaseItems.push(caseItem);
                    allRetainedItems.push(caseItem);
                  }else{
                    migratedCaseItems.push(caseItem);
                  }
                }else{
                  migratedCaseItems.push(caseItem);
                }
    
              });  
              
              console.log("After modification Retained Items are ..");
              console.log(retainedCaseItems);
              console.log("After modification allRetained Items are");
              console.log(allRetainedItems);
              retainedCaseColumns = this.props.retainedCaseColumns.slice();
              retainedCaseColumns.forEach(column => {
                if(column.fieldName === 'dueDate'){
                  column.isSortedDescending = true;
                  column.isSorted = true;
                }else{
                  column.isSorted = false;
                  column.isSortedDescending = false;
                }
              });
              retainedCaseItems = PivotTabsLargeExample._copyAndSort(retainedCaseItems, 'dueDate', true);
              console.log("After sorting Retained Items are::");
              console.log(retainedCaseItems);
              this.props.modifyState(migratedCaseItems, allRetainedItems,retainedCaseItems, true, retainedCaseColumns);


            }
            //For Retained TAB Scenario
            else if(this.props.caseType === 'Retained'){
                let normalCaseItems : any[] = [];
                let migratedCaseItems : any[] = [];
                let retainedCaseItems : any[] = [];
                let allRetainedItems : any[] = [];

                if(this.props.normalItems && this.props.normalItems.length > 0){
                  normalCaseItems = this.props.normalItems;
                }
                if(this.props.migratedItems && this.props.migratedItems.length > 0){
                  migratedCaseItems = this.props.migratedItems;
                }
                if(this.props.allRetainedItems && this.props.allRetainedItems.length > 0){
                  allRetainedItems = this.props.allRetainedItems;
                }
                console.log("Retained TAB****");
                console.log(allRetainedItems);
                this._allItems.forEach((caseItem : any) => {
                  if(caseItem.caseId === caseId){
                    console.log("Retained CaseId::"+caseItem.caseId);
                    caseItem.status = singleCaseItem[0].Status;
                    caseItem.retainedOrDeleted = singleCaseItem[0].Retained_Or_Deleted;
                    caseItem.closedDate = singleCaseItem[0].ClosedDate;
                    caseItem.closedBy = singleCaseItem[0].ClosedBy;
                    caseItem.urn = singleCaseItem[0].URN;
                    caseItem.deleteConsentDate = singleCaseItem[0].DeleteConsentDate;
                    caseItem.deleteConsentBy = singleCaseItem[0].DeleteConsentBy;
                    if(caseItem.retainedOrDeleted === 'Delete'){
                      if(caseItem.title.startsWith('mig-Case')){
                        migratedCaseItems.push(caseItem);
                      }else if(caseItem.title.startsWith('Case')){
                        normalCaseItems.push(caseItem);
                      }
                      //Removing the 'delete' item from Retained List
                      allRetainedItems = allRetainedItems.filter(item => item.caseId !== caseId);
                    }else{
                      caseItem.retainJustification = singleCaseItem[0].RetainJustification;
                      caseItem.retainConsentDate = singleCaseItem[0].RetainConsentDate;
                      caseItem.retainConsentBy = singleCaseItem[0].RetainConsentBy;
                      if(singleCaseItem[0].NextReviewDate){
                        let reviewDate = new Date(singleCaseItem[0].NextReviewDate);
                        caseItem.nextReviewDate = reviewDate;
                        caseItem.nextReviewDateString = reviewDate.toLocaleDateString();
                      }
                      
                      retainedCaseItems.push(caseItem);
                      //Removing the item and then adding it to avoid duplication in 'allRetainedItems' array
                      allRetainedItems = allRetainedItems.filter(item => item.caseId !== caseId);
                      allRetainedItems.push(caseItem);
                    }
                  }else{
                    
                    retainedCaseItems.push(caseItem);
                  }
                });
                console.log("allRetainedItems after delete consent");
                console.log(allRetainedItems);
                normalCaseItems = PivotTabsLargeExample._copyAndSort(normalCaseItems, 'dueDate', true);
                migratedCaseItems = PivotTabsLargeExample._copyAndSort(migratedCaseItems, 'dueDate', true);
                this.props.modifyState(retainedCaseItems, allRetainedItems, migratedCaseItems, normalCaseItems, true);


            }
            
          });
         });
        
        this.setState({
            ishidden : true,
            inputValue:'',
            selectedClosedDate:null
        });
      }else{
        DialogOnClose._changed = true;
        this.setState({
          valid : false
        });
      }

    }

    private _onChangeOfURN = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        console.log("new value of text box::"+newValue);
        DialogOnClose._changed = true;
        this.setState({
                inputValue : newValue,
                //selectedDate : this.state.selectedDate
        });
    }

    private _onChangeOfRetainComments = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, 
                                          newValue?: string) => {
        console.log("new value of Retain Justification::"+newValue);
        DialogOnClose._changed = true;
        this.setState({
          retainJustification : newValue
        });                                          

    }

    private _onSelectDate = (date : Date) => {
        console.log(date);
        DialogOnClose._changed = true;
        this.setState({
            selectedClosedDate : date,
            //inputValue : this.state.inputValue
        });
    }

    private _onSelectNextReviewDate = (date : Date) => {
      console.log("next Review Date ::"+date);
      DialogOnClose._changed = true;
      this.setState({
          selectedNextReviewDate : date
      });
    }

    public static getDerivedStateFromProps(nextProps, prevState){
        console.log("******getDerivedStateFromProps Close Dialog start*********");
        console.log('next props ::'+nextProps.hideDialog);
        console.log('prev state ::'+prevState.ishidden);
        console.log('close clicked =='+dialogClosed);
        console.log("DialogOnClose._changed::"+DialogOnClose._changed);
        let hide;
        let retainJustification = "";
        let nextReviewDate = null;
        if(DialogOnClose._changed === true){// this if loop is created to facilitate modification of values in popup
          retainJustification = prevState.retainJustification;
          nextReviewDate = prevState.selectedNextReviewDate;
        }
        console.log("Initial Retain Justification ::"+retainJustification);
        console.log("Initial next Review Date ::"+nextReviewDate);
        //load the existing values in popup for 'Retained TAB' on click of Consent
        if(nextProps.savedCaseData && nextProps.caseType === 'Retained' && retainJustification === "" 
            && nextReviewDate === null)
        {
          retainJustification = nextProps.savedCaseData.retainJustification;
          nextReviewDate = nextProps.savedCaseData.nextReviewDate;
          console.log("nextProps next Review Date::"+nextReviewDate);
          console.log("nextProps retain justification::"+retainJustification);
        }
        if(prevState.ishidden === true && !dialogClosed){
            hide = nextProps.hideDialog;
        }else if(DialogOnClose._changed){
            console.log("change going on ..");
            hide = false;
            
        }
        else {
            dialogClosed = false;
            hide = true;
        }
        return {
            ishidden : hide,
            retainJustification : retainJustification,
            selectedNextReviewDate : nextReviewDate
        };
    }

    public render(){
        console.log("is hidden in Close Dialog render ::"+this.state.ishidden);
        console.log(this.props.savedCaseData);
        let savedCaseData = this.props.savedCaseData;
        this._allItems = this.props.allItems;
       
       
        return (
            <Dialog
            hidden={this.state.ishidden}
            onDismiss={this._close}
            dialogContentProps={dialogContentProps}
            modalProps={modelProps}
          >
            <p></p>
            {savedCaseData && savedCaseData.isChecked === true && this._validClosedDate === false &&
              <p style={{color:'red'}}>Please enter Closed Date</p>
            }
           {savedCaseData && savedCaseData.isChecked === true && 
             <DatePicker
            className={controlClass.control}
            firstDayOfWeek={DayOfWeek.Sunday}
            strings={DayPickerStrings}
            onSelectDate={this._onSelectDate}
            placeholder="Select a Closed Date..."
            ariaLabel="Select a date"
            value={this.state.selectedClosedDate}
          />
           }
           {savedCaseData && savedCaseData.isDeleteSelected === true && this._validURN === false && 
              <p style={{color:'red'}}>Please enter URN</p>
           }
          {savedCaseData && savedCaseData.isDeleteSelected === true &&
           <TextField id="inputData" value={this.state.inputValue} onChange={this._onChangeOfURN} 
           label="Enter URN " required />
          }

          {savedCaseData && savedCaseData.isRetainSelected === true && this._validNextReviewDate === false && 
            <p style={{color:'red'}}>Please enter Next Review Date</p>
          }

          
          {savedCaseData && savedCaseData.isRetainSelected === true && this.props.caseType === 'Retained' &&
            <DatePicker className={controlClass.control} firstDayOfWeek={DayOfWeek.Sunday} 
            strings={DayPickerStrings} onSelectDate={this._onSelectNextReviewDate}
            placeholder="Select Next Review Date..." ariaLabel="Select a date"
            value={this.state.selectedNextReviewDate}/>
          }

          {savedCaseData && savedCaseData.isRetainSelected === true && this._validRetainJustification === false &&
            <p style={{color:'red'}}>Please enter Retain Justification</p>
          }
          {savedCaseData && savedCaseData.isRetainSelected === true && 
            <TextField value={this.state.retainJustification} onChange={this._onChangeOfRetainComments}
            label="Enter Retain Justification " required />
          }

            <DialogFooter>
                <PrimaryButton text="Submit" onClick={() => this._save()}/>
                <PrimaryButton text="Cancel" onClick={() => this._close()}/>
            </DialogFooter>
            
          </Dialog> 
        );
    }

}