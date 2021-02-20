import * as React from 'react';
import { IPivotStyles, Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DetailsListExample } from './ListComponents/DetailsListExample';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';

import { ISearchBoxStyles, SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { mergeStyles, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
//import { sp } from '@pnp/sp';
//import pnp from 'sp-pnp-js';
import { DatePicker, DayOfWeek, PrimaryButton } from 'office-ui-fabric-react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
//import { ICamlQuery, sp } from '@pnp/sp/presets/all';
import { IList, sp, SPRest } from "@pnp/sp-commonjs";
import { DayPickerStrings } from "./ListComponents/DialogOnClose";
//import {CSVLink} from 'react-csv';
import { CommandBarButton } from 'office-ui-fabric-react';

//sample


const pivotStyles: Partial<IPivotStyles> = {
  link: {
    width: "150px",
   
    //border: "thin solid",
    selectors:{
      ':hover':{
        backgroundColor:'#F2F2F2',
        color: 'black'
        
      }
      
    }
  }

  
  /*linkIsSelected: {
    width: "150px",
    backgroundColor: 'red',
    border: "thin solid red"
  }*/
};

const controlClass = mergeStyleSets({
  control: {
    margin: '15px',
    maxWidth: '300px',
  },
});
let searched = false;
const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width :200} };
const headerStyles = mergeStyles({
  //color:'white',
  //backgroundColor:'rgb(0,120,212)',
  border: 'thin solid'
 
}); 










export interface IPivotTabsLargeExampleProps {
  //context: WebPartContext;
  sp: SPRest;
}
export interface IPivotTabsLargeExampleState {
  normalCaseItems: any[];
  migratedCaseItems: any[];
  retainedCaseItems: any[];
  deletedCaseItems: any[];

  migratedCaseColumns : IColumn[];
  normalCaseColumns : IColumn[];
  retainedCaseColumns : IColumn[];
  deletedCaseColumns : IColumn[];

  itemsModified : boolean;
  displayFilter : boolean;

  startDate : Date;
  endDate : Date;

  allRetainedItems : any[];
  valid : boolean;
  
}
export class PivotTabsLargeExample extends React.Component<IPivotTabsLargeExampleProps,IPivotTabsLargeExampleState> {
 
  public _normalCaseItems: any[];
  public _migratedCaseItems: any[];
  public _retainedCaseItems: any[];
  public _deletedCaseItems: any[];
  
  private _normalCaseColumns: IColumn[];
  private _migratedCaseColumns: IColumn[];
  private _retainedCaseColumns: IColumn[];
  private _deletedCaseColumns: IColumn[];

  private _caseListSize : number;
  private _caseRows : any[];
  public _selectedTabItem : string;
  private _isSearched : boolean = false;
  private _isFiltered : boolean = false;

  private _currentUserName : string;
  
  private readonly sp: SPRest;

  private _allRetainedItems : any[];

  private _validStartDate : boolean;
  private _validEndDate : boolean;
  private _validDateRange : boolean;




  constructor(props: IPivotTabsLargeExampleProps){
    super(props);
    this.sp = this.props.sp;
    console.log("***********sp value::"+sp);
    /* sp.setup({
      spfxContext: this.props.context
    }); */
    this._normalCaseColumns = [
      { key: 'column1', name: 'Case ID', fieldName: 'caseId', minWidth:50, maxWidth:50,isResizable:true,
         onColumnClick: this._onColumnClick},
      { key: 'column2', name: 'Policy Number', fieldName: 'policyNumber', minWidth: 50,maxWidth:100, isResizable:true,isMultiline:true},
      { key: 'column3', name: 'Last Updated', fieldName: 'dueDate', minWidth: 100, maxWidth: 100,isResizable: true, isMultiline:true,
      isSorted:true, isSortedDescending:true,onColumnClick: this._onColumnClick, onRender: (item: any) => {
          return <span>{item.dateString}</span>;
        }},
      { key: 'column4', name: 'Case Status', fieldName: 'status', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true},
      { key: 'column5', name: 'Close Case', fieldName:'closeCase', minWidth: 50, maxWidth:100, isResizable: true},
      { key: 'column6', name: 'Case Offline Files Deleted', fieldName:'delete', minWidth: 100, maxWidth:100,isResizable: true},
      { key: 'column7', name: 'Retain', fieldName:'retain', minWidth: 50, maxWidth:50,isResizable: true},
      { key: 'column8', name: '', fieldName:'save', minWidth: 50, maxWidth:100,isResizable: true}
      //{ key: 'column9', name: 'Scenario.', fieldName:'comments', minWidth: 50, maxWidth:100, isResizable: true, isMultiline: true}
    ];

    this._migratedCaseColumns = [
      { key: 'column1', name: 'Case ID', fieldName: 'caseId', minWidth:50, maxWidth:50,isResizable:true,
      onColumnClick: this._onColumnClick, data:'number'},
      { key: 'column2', name: 'Policy Number', fieldName: 'policyNumber', minWidth: 50,maxWidth:100, isResizable:true, isMultiline:true},
      { key: 'column3', name: 'Date of GK Note', fieldName: 'dueDate', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true,
        isSorted:true, isSortedDescending:true, onColumnClick : this._onColumnClick,onRender: (item: any) => {
          return <span>{item.dateString}</span>;
        }},
      { key: 'column4', name: 'Case Status', fieldName: 'status', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true},
      { key: 'column5', name: 'Close Case', fieldName:'closeCase', minWidth: 50, maxWidth:100, isResizable: true},
      { key: 'column6', name: 'Case Offline Files Deleted', fieldName:'delete', minWidth: 100, maxWidth:100,isResizable: true},
      { key: 'column7', name: 'Retain', fieldName:'retain', minWidth: 50, maxWidth:50,isResizable: true},
      { key: 'column8', name: '', fieldName:'save', minWidth: 50, maxWidth:100, isResizable: true}
      //{ key: 'column9', name: 'Scenario.', fieldName:'comments', minWidth: 50, maxWidth:100, isResizable: true, isMultiline: true}
    ];

    this._retainedCaseColumns = [
      { key: 'column1', name: 'Case ID', fieldName: 'caseId', minWidth:50, maxWidth:50,isResizable:true,
      onColumnClick: this._onColumnClick},
      { key: 'column2', name: 'Policy Number', fieldName: 'policyNumber', minWidth: 50,maxWidth:100, isResizable:true, isMultiline:true},
      { key: 'column3', name: 'Last Updated', fieldName: 'dueDate', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true,
      isSorted:true,isSortedDescending:true,onColumnClick : this._onColumnClick,onRender: (item: any) => {
        return <span>{item.dateString}</span>;
      }},
      { key: 'column4', name: 'Case Status', fieldName: 'status', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true},
      { key: 'column5', name: 'Close Case', fieldName:'closeCase', minWidth: 50, maxWidth:100, isResizable: true},
      { key: 'column6', name: 'Case Offline Files Deleted', fieldName:'delete', minWidth: 100, maxWidth:100,isResizable: true},
      { key: 'column7', name: 'Retain', fieldName:'retain', minWidth: 50, maxWidth:50,isResizable: true},
      { key: 'column8', name: 'Comments', fieldName:'retainJustification', minWidth: 50, maxWidth:100, isResizable: true, isMultiline: true},
      { key: 'column9', name: 'Next Review Date', fieldName:'nextReviewDate',minWidth: 50, maxWidth:100, 
        isResizable: true, isMultiline: true, onColumnClick : this._onColumnClick, onRender:(item: any) => {
          return <span>{item.nextReviewDateString}</span>;
        }}, 
      { key: 'column10', name: '', fieldName:'save', minWidth: 50, maxWidth:100, isResizable: true}
     
     
    ];

    this._deletedCaseColumns = [
      { key: 'column1', name: 'Case ID', fieldName: 'caseId', minWidth:50, maxWidth:50,isResizable:true,
      onColumnClick: this._onColumnClick},
      { key: 'column2', name: 'URN', fieldName: 'urn', minWidth: 50,maxWidth:100, isResizable:true},
      { key: 'column3', name: 'Deleted Date', fieldName: 'deletedDate', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true,
      isSorted:true, isSortedDescending:true, onColumnClick : this._onColumnClick,onRender: (item: any) => {
        return <span>{item.dateString}</span>;
      }},
      { key: 'column4', name: 'Deleted By', fieldName: 'deletedBy', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true},
      { key: 'column5', name: 'Comments', fieldName:'comments', minWidth: 50, maxWidth:100, isResizable: true,isMultiline:true}
      
    ];

    
   
    this._normalCaseItems = [];
    this._migratedCaseItems = [];
    this._retainedCaseItems = [];
    this._deletedCaseItems = [];
    this._caseRows = [];
    this._currentUserName = '';
    this._allRetainedItems = [];
    this._validStartDate = true;
    this._validEndDate = true;
    this._validDateRange = true;

    
    
    

    this.state = {
      normalCaseItems : this._normalCaseItems,
      migratedCaseItems : this._migratedCaseItems,
      retainedCaseItems : this._retainedCaseItems,
      deletedCaseItems : this._deletedCaseItems,

      normalCaseColumns : this._normalCaseColumns,
      migratedCaseColumns : this._migratedCaseColumns,
      retainedCaseColumns : this._retainedCaseColumns,
      deletedCaseColumns : this._deletedCaseColumns,

      itemsModified : false,
      displayFilter : false,

      startDate : null,
      endDate : null,

      allRetainedItems : this._allRetainedItems,
      valid : true

    };

    this._onChange = this._onChange.bind(this);
    this._onColumnClick = this._onColumnClick.bind(this);
    
    
  }

  public async componentDidMount() : Promise<void>{
    
    let currentDate = new Date();
    let thresholdDate = new Date(currentDate.getFullYear()-6, currentDate.getMonth(),
                                         currentDate.getDate());
    let thresholdDateString = thresholdDate.toISOString(); 
     /* istanbul ignore next */                                   
    console.log("threshold date ::::"+thresholdDate);
     /* istanbul ignore next */ 
    console.log("Date in ISO format ::"+thresholdDateString);
    try{
    
    console.log("...Current User Info===");
    let currentUser = await sp.web.currentUser.get();
     
    this._currentUserName = currentUser.Title;
    //this._currentUserName = 'Rohit V';
    
    //check Migrated Case Retrieval starts
    console.log("user is =="+this._currentUserName);
    /*const list:IList = sp.web.lists.getByTitle('Case');
    console.log(list);*/
    const caseRecords: any[] = await sp.web.lists.getByTitle('Case').items.getAll();
    console.log("Total number of Cases ::"+caseRecords.length);
    let migratedCaseRecords : any[] = [];
    let normalCaseRecords : any[] = [];
    for(let caseRecord of caseRecords){
     const title = caseRecord.Title;
     if(title.startsWith('mig-Case')){
       migratedCaseRecords.push(caseRecord);
     }else if(title.startsWith("Case")){
       normalCaseRecords.push(caseRecord);
     }
    }
    console.log("**Normal Case Records length :::"+normalCaseRecords.length);
    
    console.log("**Migrated Case Records length :::"+migratedCaseRecords.length);
    
    //check Migrated Case Retrieval ends
     this.getNormalCaseItemsForRetention(normalCaseRecords, thresholdDate);
     this.getMigratedCaseItemsForRetention(migratedCaseRecords, thresholdDate);
     this.getDeletedCaseItems();
     
     for(let item of this._allRetainedItems){
        this._retainedCaseItems.push(item);
    }
    console.log("Retained Case Items copied..");
    
    this._retainedCaseItems = PivotTabsLargeExample._copyAndSort(this._retainedCaseItems,'dueDate',true);
    console.log(this._retainedCaseItems);
    this.setState({
        retainedCaseItems : this._retainedCaseItems,
        allRetainedItems : this._allRetainedItems
    });
    
     
     

    }catch(e){
      console.log("Exception in componentDidMount render List Items");
      console.error(e);
    }
    
  }

  
  //Get Normal Case Items For Retention
  public async getNormalCaseItemsForRetention(normalCaseRecords:any[], thresholdDate:Date){
    console.log("Normal cases length::"+normalCaseRecords.length);
    console.log("Threshold Date ::"+thresholdDate);
    try{
      for(let row of normalCaseRecords){
        if(row.CaseModifiedDate){
          let modifiedDate = new Date(row.CaseModifiedDate);
          if(modifiedDate.getTime() < thresholdDate.getTime()){
            let modifiedDateString = modifiedDate.toLocaleDateString();
            if(row.Retained_Or_Deleted !== 'Retain') {
              this._normalCaseItems.push(
                {
                  caseId: parseInt(row.ID),
                  title: row.Title,
                  status: row.Status,
                  dueDate: modifiedDate,
                  dateString: modifiedDateString,
                  policyNumber: row.PolicyNumber,
                  retainedOrDeleted: row.Retained_Or_Deleted,
                  closedDate: row.ClosedDate,
                  closedBy: row.ClosedBy,
                  urn: row.URN,
                  deleteConsentDate: row.DeleteConsentDate,
                  deleteConsentBy: row.DeleteConsentBy
                });
            }else{
              let nextReviewDate;
              let nextReviewDateString;
              if(row.NextReviewDate){
                nextReviewDate = new Date(row.NextReviewDate);
                nextReviewDateString = nextReviewDate.toLocaleDateString();
              }
              this._allRetainedItems.push(
                {
                  caseId: parseInt(row.ID),
                  title: row.Title,
                  status: row.Status,
                  dueDate: modifiedDate,
                  dateString: modifiedDateString,
                  policyNumber: row.PolicyNumber,
                  retainedOrDeleted: row.Retained_Or_Deleted,
                  closedDate: row.ClosedDate,
                  closedBy: row.ClosedBy,
                  urn: row.URN,
                  deleteConsentDate: row.DeleteConsentDate,
                  deleteConsentBy: row.DeleteConsentBy,
                  retainJustification: row.RetainJustification,
                  retainConsentDate: row.RetainConsentDate,
                  retainConsentBy: row.RetainConsentBy,
                  nextReviewDate : nextReviewDate,
                  nextReviewDateString : nextReviewDateString
                });
            }
          }
  
        }
      }
      
      
      this._normalCaseItems = PivotTabsLargeExample._copyAndSort(this._normalCaseItems, 'dueDate', true);
      //this._allRetainedItems = PivotTabsLargeExample._copyAndSort(this._allRetainedItems, 'dueDate', true);
      console.log("Normal and Retained in Get Normal Case Items for Retention...");
      console.log(this._normalCaseItems);
      console.log(this._allRetainedItems);
      
      this.setState({
        normalCaseItems : this._normalCaseItems
        
      });
    }catch(e){
      console.log("Error in getNormalCaseItemsForRetention()...");
      throw e;
    }
    
  }
  

  //Get MigratedCase Items For Retention
  public async getMigratedCaseItemsForRetention(migratedCaseRecords:any[], thresholdDate:Date)
  {
    try{
       for(let row of migratedCaseRecords){
        if(row.CaseModifiedDate){
          let modifiedDate = new Date(row.CaseModifiedDate);
          if(modifiedDate.getTime() < thresholdDate.getTime()){
            let modifiedDateString = modifiedDate.toLocaleDateString();
            if(row.Retained_Or_Deleted !== 'Retain'){
              this._migratedCaseItems.push(
                {
                  caseId: parseInt(row.ID),
                  title: row.Title,
                  status: row.Status,
                  dueDate: modifiedDate,
                  dateString: modifiedDateString,
                  policyNumber: row.PolicyNumber,
                  retainedOrDeleted: row.Retained_Or_Deleted,
                  closedDate: row.ClosedDate,
                  closedBy: row.ClosedBy,
                  urn: row.URN,
                  deleteConsentDate: row.DeleteConsentDate,
                  deleteConsentBy: row.DeleteConsentBy
                });
            }else {
              let nextReviewDate;
              let nextReviewDateString;
              if(row.NextReviewDate){
                nextReviewDate = new Date(row.NextReviewDate);
                nextReviewDateString = nextReviewDate.toLocaleDateString();
              }
              this._allRetainedItems.push(
              {
                caseId: parseInt(row.ID),
                title: row.Title,
                status: row.Status,
                dueDate: modifiedDate,
                dateString: modifiedDateString,
                policyNumber: row.PolicyNumber,
                retainedOrDeleted: row.Retained_Or_Deleted,
                closedDate: row.ClosedDate,
                closedBy: row.ClosedBy,
                urn: row.URN,
                deleteConsentDate: row.DeleteConsentDate,
                deleteConsentBy: row.DeleteConsentBy,
                retainJustification: row.RetainJustification,
                retainConsentDate: row.RetainConsentDate,
                retainConsentBy: row.RetainConsentBy,
                nextReviewDate : nextReviewDate,
                nextReviewDateString : nextReviewDateString
              });
            }

          }
        }  
       }
      
      this._migratedCaseItems = PivotTabsLargeExample._copyAndSort(this._migratedCaseItems,'dueDate',true);
     // this._allRetainedItems = PivotTabsLargeExample._copyAndSort(this._allRetainedItems,'dueDate',true);
      console.log("Migrated and Retained in Get Migrated Case Items For Retention...");
      console.log(this._migratedCaseItems);
      console.log(this._allRetainedItems);
      
      this.setState({
        migratedCaseItems : this._migratedCaseItems
        
      });
    }catch(e){
      console.log("Error in getMigratedCaseItemsForRetention()...");
      throw e;
    }
  }
  //Get Deleted Case Items
  public async getDeletedCaseItems(){
    let pageResult : any;
    let deletedCaseRows : any[] = [];
    try{
      console.log("*********Get Deleted Case Items*************");
      
      const latestItemID: any[] = await sp.web.lists.getByTitle("Retention log").items.select("ID")
                                            .orderBy("ID",false).top(1).get();
      console.log(latestItemID);
      
      let deletedItemsSize:number;
    
      latestItemID.forEach(item => {
        deletedItemsSize = item.ID;
      });
      let noOfIterations: number = Math.round(deletedItemsSize / 5000);
      let remainder : number = deletedItemsSize % 5000;
      console.log("no of Iterations ::"+noOfIterations);
      console.log("Remainder ::"+remainder);
      if(noOfIterations == 0 && remainder > 0){
        console.log("Get remainder deleted Case items");
        pageResult = await this.getPagedItemForDeletedCase();
        console.log(pageResult);
        deletedCaseRows.push(pageResult.Row);
       }else if(noOfIterations > 0){
        for(let index=0;index<noOfIterations;index++){
          if(index == 0){
            console.log("index:: 0 for Deleted Case");
            pageResult = await this.getPagedItemForDeletedCase();
            console.log(pageResult);
            deletedCaseRows.push(pageResult.Row);
            
          }else{
            if(typeof pageResult.NextHref !== 'undefined'){
              console.log("Deleted Case index::"+index);
              const pageToken = pageResult.NextHref.substring(1);
              console.log("pageToken::"+pageToken);
              pageResult = await this.getPagedItemForDeletedCase(pageToken);
              deletedCaseRows.push(pageResult.Row);
            }
            
          }
        }
        if(remainder > 0){
          if(typeof pageResult.NextHref !== 'undefined'){
            console.log("Deleted Case Iterations > 0 and Remainder > 0");
            const pageToken = pageResult.NextHref.substring(1);
            console.log("pageToken::"+pageToken);
            pageResult = await this.getPagedItemForDeletedCase(pageToken);
            this._caseRows.push(pageResult.Row);
          }
        }
       }

       console.log("Deleted caseRows length ::"+deletedCaseRows.length);
       for(let rowArr of deletedCaseRows){
         console.log("No of Rows ::"+rowArr.length);
         for(let row of rowArr){
           
           let deletedDate = new Date(row.DeletedDate);
           console.log(deletedDate);
           let deletedDateString = deletedDate.toLocaleDateString();
           /* console.log(deletedDateString);
           console.log(row.ConsentBy);
           console.log(row.Linked_Cases); */
           let caseIdString : string = row.CaseID;
           caseIdString = caseIdString.replace(/,/g,"");
           console.log(parseInt(caseIdString));
          
           this._deletedCaseItems.push({
            caseId: parseInt(caseIdString),
            urn: row.URN,
            deletedDate: deletedDate,
            dateString: deletedDateString,
            deletedBy: row.ConsentBy
          });
           console.log("----------------------");
         }
       }
      
      this.setState({
        deletedCaseItems : this._deletedCaseItems
      });
    }catch(e){
      console.log("Error in getDeletedCaseItems for Cases ..");
      console.log(e);
      throw e;
    }


  }

  
  private async getPagedItemForDeletedCase(pageToken?: string){
    console.log("Get Paged Items for Deleted Cases..");
    const list:IList = sp.web.lists.getByTitle("Retention log");
    return list.renderListDataAsStream({
      ViewXml: `<View>
      <Query>
        <OrderBy><FieldRef Name='DeletedDate' Ascending='FALSE'/></OrderBy>
      </Query>
      <ViewFields>
        <FieldRef Name='CaseID' />
        <FieldRef Name='URN' />
        <FieldRef Name='DeletedDate' />
        <FieldRef Name='ConsentBy' />
        <FieldRef Name='Linked_Cases' />
      </ViewFields>
      </View>`,
      Paging: pageToken
    });
  }

   public _onChange(event?: React.ChangeEvent<HTMLInputElement>, newValue?: string){
   console.log("***new value search _onchange::"+newValue);
   this._isSearched = true;
   if(this.state.itemsModified === true){
     console.log("Items modified");
     this._normalCaseItems = this.state.normalCaseItems;
     this._migratedCaseItems = this.state.migratedCaseItems;
     this._retainedCaseItems = this.state.retainedCaseItems;
   }
   
   if(this._selectedTabItem === 'Normal Cases'){
    this.setState({
      normalCaseItems: newValue ? this._normalCaseItems
                          .filter(i => i.caseId.toString().includes(newValue) || (i.policyNumber !== null && i.policyNumber.toString().includes(newValue))) : this._normalCaseItems,
      itemsModified : false
      
    });
   }else if(this._selectedTabItem === 'Migrated Cases'){
     this.setState({
      migratedCaseItems: newValue ? this._migratedCaseItems
                        .filter(i => i.caseId.toString().includes(newValue) || (i.policyNumber !== null && i.policyNumber.toString().includes(newValue))) : this._migratedCaseItems,
      itemsModified : false                        
     
     });
   }else if(this._selectedTabItem === 'Retained Cases'){
      
      this.setState({
      retainedCaseItems: newValue ? this._retainedCaseItems
                        .filter(i => i.caseId.toString().includes(newValue) || (i.policyNumber !== null && i.policyNumber.toString().includes(newValue))) : this._retainedCaseItems,
      itemsModified : false                        
      
     });
   }else if(this._selectedTabItem === 'Deleted Cases'){
     this.setState({
      deletedCaseItems: newValue ? this._deletedCaseItems
                        .filter(i => i.caseId.toString().includes(newValue)) : this._deletedCaseItems
     });
   }else {
     console.log("Normal Case search::");
    this.setState({
      normalCaseItems: newValue ? this._normalCaseItems
                          .filter(i =>i.caseId.toString().includes(newValue) || (i.policyNumber !== null && i.policyNumber.toString().includes(newValue))) : this._normalCaseItems,
      itemsModified : false                          
      
    });
   }
  
   
   
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        let newItems : any[] = [];
        let newColumns : IColumn[] =[];
       
        if(this._selectedTabItem === 'Migrated Cases'){
          const{ migratedCaseColumns, migratedCaseItems } = this.state;
          console.log("Migrated case Items initial length is ::"+migratedCaseItems.length);
          newItems = migratedCaseItems;
          newColumns = migratedCaseColumns.slice();
        }else if(this._selectedTabItem === 'Normal Cases'){
          const{ normalCaseColumns, normalCaseItems } = this.state;
          newItems = normalCaseItems;
          newColumns = normalCaseColumns.slice();
        }else if(this._selectedTabItem === 'Retained Cases'){
          const{ retainedCaseColumns, retainedCaseItems } = this.state;
          newItems = retainedCaseItems;
          newColumns = retainedCaseColumns.slice();
        }else if(this._selectedTabItem === 'Deleted Cases'){
          const{ deletedCaseColumns, deletedCaseItems } = this.state;
          newItems = deletedCaseItems;
          newColumns = deletedCaseColumns.slice();
        }else {
          console.log("Normal Case column clicked...");
          const{ normalCaseColumns, normalCaseItems } = this.state;
          newItems = normalCaseItems;
          newColumns = normalCaseColumns.slice();
        }
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
       
        console.log("...column clicked is "+column.fieldName);
        console.log("typeof field ::"+typeof column.fieldName);
        console.log("Items old length ::"+newItems.length);
        console.log(newItems);
        newColumns.forEach((newCol: IColumn) => {
          if (newCol === currColumn) {
            currColumn.isSortedDescending = !currColumn.isSortedDescending;
            currColumn.isSorted = true;
          } else {
            newCol.isSorted = false;
            newCol.isSortedDescending = true;
          }
        });
        if(column.fieldName === 'nextReviewDate'){
          console.log("Filter all undefined Item values");
          let itemsWithReviewDate : any[] = newItems.filter(item => item.nextReviewDate !== undefined);
          let itemsWithoutReviewDate : any[] = newItems.filter(item => item.nextReviewDate === undefined);
          newItems = PivotTabsLargeExample._copyAndSort(itemsWithReviewDate, column.fieldName,
                                    currColumn.isSortedDescending);
          for(let item of itemsWithoutReviewDate){
            newItems.push(item);
          }                                    
        }else{
          newItems = PivotTabsLargeExample._copyAndSort(newItems, column.fieldName, 
                                currColumn.isSortedDescending);
        }
        
        console.log("New Items after sorting");
        console.log(newItems);                              
        if(this._selectedTabItem === 'Migrated Cases'){
          this.setState({
            migratedCaseItems : newItems,
            migratedCaseColumns : newColumns
          });
        } else if(this._selectedTabItem === "Normal Cases"){
          this.setState({
            normalCaseItems : newItems,
            normalCaseColumns : newColumns
          });
        } else if(this._selectedTabItem === "Retained Cases"){
          this.setState({
            retainedCaseItems : newItems,
            retainedCaseColumns : newColumns
          });
        } else if(this._selectedTabItem === "Deleted Cases"){
          this.setState({
            deletedCaseItems : newItems,
            deletedCaseColumns : newColumns
          });
        } else{
          this.setState({
            normalCaseItems : newItems,
            normalCaseColumns : newColumns
          });
        }                         
       
  }

  public static _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    console.log("key = "+key);
    //return items.slice(0).sort( (a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    return items.slice(0).sort( (a:T, b:T) => {
              
      return (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1;
    });
  }
  

  

  

  private _onLinkClick = (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>):void => {
    this._selectedTabItem = item.props.headerText;
    console.log("Pivot item selected is =::;"+this._selectedTabItem);
    let displayFilter : boolean;
    if(this._selectedTabItem === 'Retained Cases'){
      displayFilter = true;
    }else{
      displayFilter = false;
    }
    this.setState({
      displayFilter : displayFilter
    });
   

  }

  public _filter = () : void => {
    console.log("Filter based on provided Date Range...");
    //allRetainedItems contains all the Retained Items present in Case List and is not modified based
    //on filter parameters
    console.log(this.state.allRetainedItems);
    if(!this.state.startDate || this.state.startDate === null){
      this._validStartDate = false;
    }else{
      this._validStartDate = true;
    }

    if(!this.state.endDate || this.state.endDate === null){
      this._validEndDate = false;
    }else {
      this._validEndDate = true;
    }

    if(this._validStartDate === true && this._validEndDate === true && 
        (this.state.startDate > this.state.endDate)){
          this._validDateRange = false;
        } else {
          this._validDateRange = true;
        }
    if(this._validStartDate === true && this._validEndDate === true && this._validDateRange === true){
      //this._retainedCaseItems = [];
      this._isFiltered = true;
      this._retainedCaseItems = this.state.allRetainedItems.filter(record => {
      if(record.nextReviewDate){
        let nextReviewDate = new Date(record.nextReviewDate);
        if(nextReviewDate.getTime() >= this.state.startDate.getTime() && 
          nextReviewDate.getTime() <= this.state.endDate.getTime()){
          return true;
        }
          return false;
      } /* else{//newly Retained Items coming from Normal or Migrated TAB are also included
        return true;
      } */
    });
    let retainedCaseColumns: IColumn[] = this.state.retainedCaseColumns.slice();
    retainedCaseColumns.forEach(column => {
        if(column.fieldName === 'nextReviewDate'){
          column.isSorted = true;
          column.isSortedDescending = false;
        }else{
          column.isSorted = false;
          column.isSortedDescending = true;
        }
    });
    this._retainedCaseItems = PivotTabsLargeExample._copyAndSort(this._retainedCaseItems, 
      'nextReviewDate', false);
    console.log(this._retainedCaseItems);
    this.setState({
      retainedCaseItems : this._retainedCaseItems,
      retainedCaseColumns : retainedCaseColumns
      
    });
  }else{
    this.setState({
      valid : false
    });
  }

  }

  public _onSelectStartDate = (date : Date) => {
    this.setState({
        startDate : date
    });
  }

  public _onSelectEndDate = (date : Date) => {
    this.setState({
      endDate : date
    });
  }

  

 

  public  render() {
    
   
    console.log("3..normal Case Items state length.."+this.state.normalCaseItems.length);
    console.log("3.Retained Cases length::"+this.state.retainedCaseItems.length);
    console.log(this.state.retainedCaseItems);

    
    

    return (
      
      <div id='parent'>

         {this.state.displayFilter === true && this._validStartDate === false &&
          <span style={{color:'red'}}>Please Enter Start Date </span>
         }

         {this.state.displayFilter === true && this._validEndDate === false && 
          <span style={{color:'red'}}>Please Enter End Date</span>
         }

         {this.state.displayFilter === true && this._validDateRange === false && 
          <span style={{color:'red'}}>Please ensure End Date is not lesser than Start Date</span>
         }
        
         {this.state.displayFilter === true && 
         
         <div style={{display : "flex"}}>
          <DatePicker
          className={controlClass.control}
          firstDayOfWeek={DayOfWeek.Sunday}
          strings={DayPickerStrings}
          onSelectDate={this._onSelectStartDate}
          placeholder="Select Start Date..."
          ariaLabel="Select a date"
          value={this.state.startDate}
         />

          <DatePicker className={controlClass.control}
           firstDayOfWeek = {DayOfWeek.Sunday}
           strings={DayPickerStrings}
           onSelectDate={this._onSelectEndDate}
           placeholder="Select End Date..."
           ariaLabel="Select a date"
           value={this.state.endDate} />

           <PrimaryButton className={controlClass.control} onClick={() => this._filter()} text="Filter"/>

         </div>

        }
        <div style={{ display: "flex", justifyContent: "flex-end" }}>
          {/* {this._selectedTabItem !== 'Migrated Cases' && this._selectedTabItem !== 'Retained Cases' && this._selectedTabItem !== 'Deleted Cases' && 
          <CSVLink style={{margin:'15px'}} data={this.state.normalCaseItems} filename={'Normal_Cases.csv'}>
            <CommandBarButton  iconProps={{ iconName: 'ExcelLogoInverse' }} text='Export Normal Cases to Excel' />
          </CSVLink>
          }

          {this._selectedTabItem === 'Migrated Cases' &&
          <CSVLink style={{margin:'15px'}} data={this.state.migratedCaseItems} filename={'Migrated_Cases.csv'}>
            <CommandBarButton  iconProps={{ iconName: 'ExcelLogoInverse' }} text='Export Migrated Cases to Excel' />
          </CSVLink>
          }

          {this._selectedTabItem === 'Retained Cases' &&
          <CSVLink style={{margin:'15px'}} data={this.state.retainedCaseItems} filename={'Retained_Cases.csv'}>
            <CommandBarButton  iconProps={{ iconName: 'ExcelLogoInverse' }} text='Export Retained Cases to Excel' />
          </CSVLink>
          }

          {this._selectedTabItem === 'Deleted Cases' && 
          <CSVLink style={{margin:'15px'}} data={this.state.deletedCaseItems} headers={this._deleteHeaders} filename={'Deleted_Cases.csv'}>
            <CommandBarButton iconProps={{iconName: 'ExcelLogoInverse'}} text="Export Deleted Cases to Excel" />
          </CSVLink>
          } */}

        <SearchBox id="searchbox" styles={searchBoxStyles} placeholder="Search" 
                    onChange={this._onChange} />
          
        </div> 
        <Pivot id="pivotelement" styles={pivotStyles} onLinkClick={this._onLinkClick}
          aria-label="Links of Large Tabs Pivot Example"
          linkFormat={PivotLinkFormat.tabs}
          linkSize={PivotLinkSize.large}
        >
          <PivotItem id="normalCasePivotItem" headerText="Normal Cases">
            <p></p>
            <DetailsListExample modifyState = {(normalItems, allRetainedItems, retainedItems, itemsModified, retainedCaseColumns) => 
                                                this.setState({normalCaseItems : normalItems,
                                                              allRetainedItems : allRetainedItems,
                                                              retainedCaseItems : retainedItems,
                                                              itemsModified : itemsModified,
                                                            retainedCaseColumns : retainedCaseColumns})} 
                                
                                items={this.state.normalCaseItems} 
                                retainedItems={this.state.retainedCaseItems}
                                allRetainedItems = {this.state.allRetainedItems} 
                                columns={this.state.normalCaseColumns}
                                retainedCaseColumns = {this.state.retainedCaseColumns}
                                searched={this._isSearched}  caseType='Normal' 
                                currentUserName={this._currentUserName}/>
          </PivotItem>
          <PivotItem id="migratedCasePivotItem" headerText="Migrated Cases">
            <p></p>
            <DetailsListExample modifyState = {(migratedItems, allRetainedItems, retainedItems,itemsModified, retainedCaseColumns) => 
                                                  this.setState({migratedCaseItems : migratedItems, 
                                                                allRetainedItems : allRetainedItems,
                                                                retainedCaseItems : retainedItems,
                                                                itemsModified : itemsModified,
                                                                retainedCaseColumns : retainedCaseColumns})} 
                                items={this.state.migratedCaseItems} 
                                columns={this.state.migratedCaseColumns}
                                retainedCaseColumns = {this.state.retainedCaseColumns}
                                retainedItems={this.state.retainedCaseItems}
                                allRetainedItems = {this.state.allRetainedItems} 
                                searched={this._isSearched} caseType='Migrated' 
                                currentUserName={this._currentUserName} />
            
          </PivotItem>
          <PivotItem id="retainedCasePivotItem" headerText="Retained Cases">
            <p></p>
            <DetailsListExample modifyState = {(retainedItems, allRetainedItems, migratedItems, normalItems, itemsModified) => 
                                                  this.setState({retainedCaseItems : retainedItems,
                                                                allRetainedItems : allRetainedItems, 
                                                                migratedCaseItems : migratedItems, 
                                                                normalCaseItems : normalItems,
                                                                itemsModified : itemsModified
                                                              })} 
                                items={this.state.retainedCaseItems} columns={this.state.retainedCaseColumns} 
                                searched={this._isSearched} caseType='Retained'
                                normalItems={this.state.normalCaseItems}
                                migratedItems={this.state.migratedCaseItems}
                                allRetainedItems={this.state.allRetainedItems}
                                filtered={this._isFiltered} 
                                currentUserName={this._currentUserName} />
            
          </PivotItem>
          <PivotItem id="deletedCasePivotItem" headerText="Deleted Cases">
            <p></p>
            <DetailsListExample modifyState = {(items) => this.setState({deletedCaseItems : items})} 
                                items={this.state.deletedCaseItems} columns={this.state.deletedCaseColumns} 
                                searched={this._isSearched} caseType='Deleted' currentUserName={this._currentUserName}/>
              
          </PivotItem>
        </Pivot>
      </div>
    );
  }

  
}

