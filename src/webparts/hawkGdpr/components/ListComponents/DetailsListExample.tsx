import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, CheckboxVisibility, IDetailsRowProps, DetailsHeader, IDetailsHeaderProps, IDetailsHeaderStyles, DetailsRow, IDetailsRowStyles } from 'office-ui-fabric-react/lib/DetailsList';

import { Fabric } from 'office-ui-fabric-react/lib/Fabric';

import { Checkbox, IChoiceGroupOption, mergeStyles, PrimaryButton } from 'office-ui-fabric-react';
import Pagination from 'office-ui-fabric-react-pagination';


import { DialogOnClose } from './DialogOnClose';











export interface IDetailsListExampleProps {
    items: any[];
    columns: IColumn[];
    caseType : string;
    searched : boolean;
    modifyState : Function;
    currentUserName : string;
    retainedItems? : any[];
    normalItems?: any[];
    migratedItems?: any[];
    allRetainedItems? : any[];
    retainedCaseColumns? : IColumn[];
    filtered? : boolean;
   
}

export interface IDetailsListExampleState {
  currentPage : number;
  pageSize : number;
  //hideDialogRetain : boolean;
  hideDialogClose : boolean;
  rowKey : number;
  isDeleteChecked : boolean;
  isRetainChecked : boolean;

 
  
 
}


let checkBoxData = new Map();
let savedCaseData = new Map();
let deleteSelectedData = new Map();
let retainSelectedData = new Map();


const PAGE_SIZE = 10;
const radiobuttonStyles = mergeStyles({
  //color:'white',
  //backgroundColor:'rgb(0,120,212)',
  border: '0px',
  width: '100%',
  height: '1.5em'
 
}); 
/* const listStyles : Partial<IDetailsListStyles> = {
  root:{border: 'thin solid blue'},
  
 }; */
 const deleteOption: IChoiceGroupOption[] = [
  { key: 'Delete',text: '' }
  
];
 const retainOption : IChoiceGroupOption[] = [
   {key : 'Retain', text:''}
 ];
 const detailsHeaderStyles : Partial<IDetailsHeaderStyles> = {
   root:{
     padding : '0px',
     selectors:{
        '.ms-DetailsHeader-cellName':{
           whiteSpace:'normal',
           textOverflow:'clip',
           lineHeight:'normal',
           width:'100px',
           textAlign: 'center'
          
        }
        
     }
    }
   
 };

 const detailsRowStyles : Partial<IDetailsRowStyles> = {
  root:{
    selectors:{
      '.ms-DetailsRow-cell':{
        border: 'thin solid'
      }
    }
  }
 };

export class DetailsListExample extends React.Component<IDetailsListExampleProps, IDetailsListExampleState>{
    private _filteredItems: any[];
    private static _numberOfPages: number;
    private _endIndex : number;
    private _startIndex : number;
    private static _saveClicked : boolean;
    
    private static _pageChanged : boolean;
    private _allItems : any[];
    


    constructor(props: IDetailsListExampleProps){
        super(props);
        this._onRenderDetailsHeader = this._onRenderDetailsHeader.bind(this);
        DetailsListExample._saveClicked = false;
        //DetailsListExample._saveAndRetainClicked = false;
        this._allItems = this.props.items;
        console.log("In constructor all Items::");
        console.log(this._allItems);
        this.state={
          currentPage : 1,
          pageSize : PAGE_SIZE,
         // hideDialogRetain : true,
          hideDialogClose : true,
          rowKey : 0,
          isDeleteChecked : false,
          isRetainChecked : false
         
        };
        DetailsListExample._numberOfPages = Math.ceil(this.props.items.length/this.state.pageSize);
        console.log("number of pages =="+DetailsListExample._numberOfPages);
       
    }

    public static getDerivedStateFromProps(nextProps, prevState){
      console.log("********getDerivedStateFromProps start*************");
      console.log(nextProps.items.length);
      DetailsListExample._numberOfPages = Math.ceil(nextProps.items.length/prevState.pageSize);
      console.log("no of pages =="+DetailsListExample._numberOfPages);
      let pageNumber = prevState.currentPage;
      console.log("prevState currentPage =="+pageNumber);
      console.log("props searched is ::"+nextProps.searched);
      console.log("page changed is ::"+DetailsListExample._pageChanged);
      console.log("TAB selected is ::"+nextProps.caseType);
      console.log("props filtered is ::"+nextProps.filtered);
      if(nextProps.caseType === 'Retained'){// created this if condition for specifically 'Retain' TAB since selection of Radio buttons change the state,but screen must stay in current page
        if(prevState.isDeleteChecked === true || prevState.isRetainChecked === true){
          console.log("Retain TAB reloading state but staying the current page.");
          pageNumber = prevState.currentPage;
        }else if(DetailsListExample._numberOfPages > 0 && pageNumber === 0){
          pageNumber = 1;
        }else if(DetailsListExample._numberOfPages > 0 && nextProps.searched && !DetailsListExample._pageChanged){
          pageNumber = 1;
        }else if(DetailsListExample._numberOfPages > 0 && nextProps.filtered && !DetailsListExample._pageChanged){
          pageNumber = 1;
        }else if(DetailsListExample._numberOfPages === 0){
          pageNumber = 0;
        }
      }
      else{
        if(DetailsListExample._numberOfPages > 0 && pageNumber === 0){
          pageNumber = 1;
        }else if(DetailsListExample._numberOfPages > 0 && nextProps.searched && !DetailsListExample._pageChanged){
          pageNumber = 1;
        }
        else if(DetailsListExample._numberOfPages === 0){
          pageNumber = 0;
        }
      }
      let isSaveClicked = DetailsListExample._saveClicked;
      console.log('save button clicked is... ::'+isSaveClicked);

      

      DetailsListExample._saveClicked = false;
      //DetailsListExample._saveAndRetainClicked = false;
      DetailsListExample._pageChanged = false;
      console.log("********getDerivedStateFromProps end*************");
      return {
        currentPage : pageNumber,
        //hideDialogRetain : !isSaveAndRetainClicked,
        hideDialogClose : !isSaveClicked,
        isDeleteChecked : false,
        isRetainChecked : false
      };
      
    }
    
    

    public render(): JSX.Element {
        this._filteredItems = [];
        
        console.log('current page =='+this.state.currentPage);
        let countOfItems = this.props.items.length;
        console.log("count of items ::"+countOfItems);

        if(this.state.currentPage === 1){
          this._startIndex = 0;
          this._endIndex = (countOfItems < PAGE_SIZE) ? --countOfItems : (PAGE_SIZE - 1);
        }else if(this.state.currentPage === DetailsListExample._numberOfPages){
          //this._startIndex = 0;
          this._endIndex = (countOfItems < PAGE_SIZE) ? --countOfItems : (this.props.items.length - 1);
        }else {
          this._endIndex = (countOfItems < PAGE_SIZE ) ? this._startIndex + --countOfItems : this._startIndex + (PAGE_SIZE - 1);
        }

        
        
        console.log('start Index ::'+this._startIndex);
        console.log('end Index ::'+this._endIndex);
        for (let i = this._startIndex; i <= this._endIndex; i++){
         if(this.props.caseType === 'Deleted'){
          this._filteredItems.push({
            key: this.props.items[i].caseId,
            caseId: this.props.items[i].caseId,
            urn: this.props.items[i].urn,
            deletedBy: this.props.items[i].deletedBy,
            deletedDate: this.props.items[i].deletedDate,
            dateString:this.props.items[i].dateString,
            comments: this.props.items[i].comments
          });
         } else{
          this._filteredItems.push({
            key: this.props.items[i].caseId,
            caseId: this.props.items[i].caseId,
            status: this.props.items[i].status,
            policyNumber: this.props.items[i].policyNumber,
            dueDate: this.props.items[i].dueDate,
            dateString: this.props.items[i].dateString,
            retainedOrDeleted: this.props.items[i].retainedOrDeleted,
            closedDate: this.props.items[i].closedDate,
            closedBy: this.props.items[i].closedBy,
            urn: this.props.items[i].urn,
            deleteConsentDate: this.props.items[i].deleteConsentDate,
            deleteConsentBy: this.props.items[i].deleteConsentBy,
            retainJustification: this.props.items[i].retainJustification,
            retainConsentDate: this.props.items[i].retainConsentDate,
            retainConsentBy : this.props.items[i].retainConsentBy,
            nextReviewDate : this.props.items[i].nextReviewDate,
            nextReviewDateString : this.props.items[i].nextReviewDateString
          });
         }
         
         
        }
       
        console.log("Filtered items count ::"+this._filteredItems.length);
       
       
                              
       
       
        return(
            <Fabric>
                <DetailsList
                compact={true}
            items={this._filteredItems}
            columns={this.props.columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            checkboxVisibility={CheckboxVisibility.hidden}
            onRenderItemColumn={this._onRenderItemColumn}
            onRenderDetailsHeader={this._onRenderDetailsHeader}
            onRenderRow={this._onRenderRow}
            
          />
          <Pagination
            currentPage={this.state.currentPage}
            totalPages={DetailsListExample._numberOfPages}
            onChange={this._onChangePage}
        />
          
          {/* <DialogExample hideDialog={this.state.hideDialogRetain} savedCaseData={savedCaseData.get(this.state.rowKey)}/>  */}
          <DialogOnClose 
            allItems={this.props.items} 
            modifyState={this.props.modifyState} 
            hideDialog={this.state.hideDialogClose} 
            savedCaseData={savedCaseData.get(this.state.rowKey)}
            currentUserName={this.props.currentUserName} 
            caseType={this.props.caseType}
            allRetainedItems={this.props.allRetainedItems}
            retainedItems={this.props.retainedItems}
            normalItems={this.props.normalItems}
            migratedItems={this.props.migratedItems}
            retainedCaseColumns={this.props.retainedCaseColumns}/>
            </Fabric>
        );
    }

    private _onRenderItemColumn = (item: any, index: number, column: IColumn): React.ReactNode => {
        //console.log(column.fieldName);
        const value = item['caseId'];
        const radiobuttonName = `radiobuttons_${value}`;
       
         if(column.fieldName === 'closeCase'){
           if(item['status'] === 'Closed'){
             //console.log("status closed checkbox disabled for caseId ::"+value);
            return <Checkbox checked={true} disabled={true} />;//Render checkbox as checked and disabled for closed Cases
           }
           return <Checkbox onChange={(ev: any, checked: boolean) => this. _onCheckboxChange(ev, checked, value)}/>;
    
         } else if(column.fieldName === 'delete'){
            if(item['retainedOrDeleted'] === 'Delete'){
              return <input type="radio" checked={true} disabled={true} className={radiobuttonStyles} name={radiobuttonName} onChange={(ev: any) => this._onDeletionClick(ev,value)}/>;
            }else if(item['retainedOrDeleted'] === 'Retain'){
              //if condition to handle click on Retain radio button in Retained TAB
              if(deleteSelectedData.get(value) === false && retainSelectedData.get(value) === true){
                //console.log("Retain clicked in Retain TAB for Case :"+value);
                return <input type="radio" checked={false} className={radiobuttonStyles} onChange={(ev: any) => this._onDeletionClick(ev, value)}/>;
              }else if(deleteSelectedData.get(value) === true && retainSelectedData.get(value) === false){
               // console.log("Delete clicked in Retain TAB for Case :"+value);
                return <input type="radio" checked={true} className={radiobuttonStyles} onChange={(ev:any) => this._onDeletionClick(ev, value)}/>;
              }
            }
            
            return <input type="radio" className={radiobuttonStyles} name={radiobuttonName} onChange={(ev: any) => this._onDeletionClick(ev,value)}/>;
         } else if(column.fieldName === 'retain'){
            if(item['retainedOrDeleted'] === 'Delete'){
              return <input type="radio" disabled={true} className={radiobuttonStyles}/>;
            }else if(item['retainedOrDeleted'] === 'Retain'){//item ['retainedOrDeleted'] === Retain condition will not succeed for Normal and Migrated TABS
              //if condition to handle click on delete radio button in Retained TAB
              if(deleteSelectedData.get(value) === true && retainSelectedData.get(value) === false){
                  //console.log("delete clicked in Retain TAB for Case :"+value);
                  return <input type="radio" checked={false} className={radiobuttonStyles} onChange={(ev: any) => this._onRetainClick(ev,value)}/>;
              }else if(deleteSelectedData.get(value) === false && retainSelectedData.get(value) === true){
                  //console.log("retain clicked in Retain TAB for Case :"+value);
                  return <input type="radio" checked={true} className={radiobuttonStyles} onChange={(ev: any) => this._onRetainClick(ev, value)}/>;
              }
              //else condition to set the below Map values on Retained TAB Initial load
              else{
                deleteSelectedData.set(value, false);
                retainSelectedData.set(value, true);
              }
              
              return <input type="radio" checked={true} className={radiobuttonStyles} onChange={(ev: any) => this._onRetainClick(ev,value)}/>;
            }
          return <input type="radio" className={radiobuttonStyles} name={radiobuttonName} onChange={(ev: any) => this._onRetainClick(ev,value)}/>;
         }
         else if(column.fieldName === 'save'){
           if(item['retainedOrDeleted'] === 'Delete') {
             return <PrimaryButton text="Consent" disabled={true} />;
           }
          return <PrimaryButton text="Consent"  onClick={() => this._saveRowData(item, column)}/>;
         }
         return item[column.fieldName];
      }

     private _onRenderRow = (detailsRowProps: IDetailsRowProps) => {
       
        return(
            //<DetailsRow styles={detailsRowStyles} {...detailsRowProps} />
            <DetailsRow {...detailsRowProps} />
        );
     }

      private _saveRowData = (item: any, column: IColumn): void => {
        
        let caseId = item['caseId'];
        let allCaseItems : any[] = [];
        let status : string;
        let retainOrDelete : string;
       
        console.log('save clicked');
        console.log("caseId is :::"+caseId);
        console.log("item retained or deleted ::"+item['retainedOrDeleted']);
        savedCaseData.set(caseId, {caseId: caseId, policyNumber: item['policyNumber'], dueDate: item['dueDate'], 
                                   status:item['status'],dateString:item['dateString'],
                                   retainedOrDeleted:item['retainedOrDeleted'],
                                   closedDate:item['closedDate'],
                                   closedBy:item['closedBy'],
                                   urn: item['urn'],
                                   deleteConsentDate: item['deleteConsentDate'],
                                   deleteConsentBy: item['deleteConsentBy'],
                                   retainJustification: item['retainJustification'],
                                   retainConsentDate: item['retainConsentDate'],
                                   retainConsentBy: item['retainConsentBy'],
                                   nextReviewDate : item['nextReviewDate'],
                                   nextReviewDateString : item['nextReviewDateString'],
                                   isChecked : checkBoxData.get(caseId), 
                                   isDeleteSelected :deleteSelectedData.get(caseId),
                                   isRetainSelected: retainSelectedData.get(caseId)} );
        console.log(savedCaseData);
        if(savedCaseData.get(caseId).isChecked === true || savedCaseData.get(caseId).isDeleteSelected === true ||
            savedCaseData.get(caseId).isRetainSelected === true){
            DetailsListExample._saveClicked = true;
            this.setState({
              hideDialogClose : false,
              rowKey : caseId
            });
        }else{
          alert("Please select any of the Checkbox or Radio button options");
        }
      }

      private _onCheckboxChange = (ev: any, checked: boolean, value: string) => {
        console.log("The checkbox is checked? ::"+checked+" and caseId is ::"+value);
        checkBoxData.set(value, checked);
        console.log(checkBoxData);
      }

      private _onDeletionClick = (ev: any, value: string) => {
        console.log("deletion radio button clicked .."+ev.currentTarget.value);
       
          if(ev.currentTarget.value === 'on'){
          if(this.props.caseType === 'Retained'){
            console.log("Retained TAB");
            deleteSelectedData.set(value, true);
            retainSelectedData.set(value, false);
            this.setState({
              isDeleteChecked : true
            });

          }
          else{
            deleteSelectedData.set(value, true);
            retainSelectedData.set(value, false);
          }
        }  
        
        

      }

      private _onRetainClick = (ev: React.FormEvent<HTMLInputElement>, value: string) => {
        console.log("retain radio button checked ::"+ev.currentTarget.value);
        if(ev.currentTarget.value === 'on'){
          if(this.props.caseType === 'Retained'){
            console.log("Retained TAB");
            retainSelectedData.set(value, true);
            deleteSelectedData.set(value, false);
            this.setState({
              isRetainChecked : true
            });
          }
          else{
          retainSelectedData.set(value, true);
          deleteSelectedData.set(value, false);
          }
        }
       
      }

      private _onRenderDetailsHeader(detailsHeaderProps: IDetailsHeaderProps) {
        return (
          <DetailsHeader styles={detailsHeaderStyles}
            {...detailsHeaderProps}
            //onRenderColumnHeaderTooltip={this.renderCustomHeaderTooltip}
          />
        );
      }

      private _onChangePage = (pageNumber : number) => {
        DetailsListExample._pageChanged = true;
       this._startIndex = (pageNumber - 1)* this.state.pageSize;
       this.setState({ 
         currentPage : pageNumber
        
        });
      }

      

      

      

      
}