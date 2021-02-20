import * as React from 'react';
import { configure, mount, ReactWrapper, ShallowWrapper, shallow } from 'enzyme';
import * as Adapter from 'enzyme-adapter-react-16';
import { IPivotTabsLargeExampleProps } from '../components/PivotLargeTabsExample';
import { IPivotTabsLargeExampleState } from '../components/PivotLargeTabsExample';
import { PivotTabsLargeExample } from '../components/PivotLargeTabsExample';
import * as sinon from 'sinon';
import { IList, sp, SPRest } from "@pnp/sp-commonjs";
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { IColumn } from 'office-ui-fabric-react';

let currentDate = new Date();
    let thresholdDate = new Date(currentDate.getFullYear()-6, currentDate.getMonth(),
                                         currentDate.getDate());
    let modifiedDate = new Date(2014,12,23);                                 
    let caseRecords : any[] = [];
    let normalCaseItem : any = {caseId: 13831,
      closedBy: null,
      closedDate: null,
      dateString: "06/01/2015",
      deleteConsentBy: null,
      deleteConsentDate: null,
      dueDate: new Date(2015, 6, 23),
      nextReviewDate: new Date(2021, 4, 21),
      nextReviewDateString: "21/04/2021",
      policyNumber: 35790,
      retainConsentBy: "Rohit V",
      retainConsentDate: new Date(2021,2,9),
      retainJustification: "Added Review Date for 13831",
      retainedOrDeleted: "Retain",
      status: "Closed",
      title: "Case:13831",
      urn: null};

    let migratedCaseItem : any = {caseId: 3783,
      closedBy: null,
      closedDate: null,
      dateString: "06/01/2015",
      deleteConsentBy: null,
      deleteConsentDate: null,
      dueDate: new Date(2015, 6, 23),
      nextReviewDate: new Date(2021, 4, 21),
      nextReviewDateString: "21/04/2021",
      policyNumber: 86239,
      retainConsentBy: "Rohit V",
      retainConsentDate: new Date(2021,2,9),
      retainJustification: "Added Review Date for 13831",
      retainedOrDeleted: "Retain",
      status: "Closed",
      title: "Case:13831",
      urn: null};

    let deletedCaseItem : any = {
      caseId : 4678,
      urn : '0544-578',
      deletedDate : new Date(2020,11,24),
      deletedBy : 'Rohit V'

    };

    let normalCaseColumns : IColumn[] = [
      { key: 'column1', name: 'Case ID', fieldName: 'caseId', minWidth:50, maxWidth:50,isResizable:true},
      { key: 'column2', name: 'Policy Number', fieldName: 'policyNumber', minWidth: 50,maxWidth:100, isResizable:true,isMultiline:true},
      { key: 'column3', name: 'Last Updated', fieldName: 'dueDate', minWidth: 100, maxWidth: 100,isResizable: true, isMultiline:true,
      isSorted:true, isSortedDescending:true},
      { key: 'column4', name: 'Case Status', fieldName: 'status', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true},
      { key: 'column5', name: 'Close Case', fieldName:'closeCase', minWidth: 50, maxWidth:100, isResizable: true},
      { key: 'column6', name: 'Case Offline Files Deleted', fieldName:'delete', minWidth: 100, maxWidth:100,isResizable: true},
      { key: 'column7', name: 'Retain', fieldName:'retain', minWidth: 50, maxWidth:50,isResizable: true},
      { key: 'column8', name: '', fieldName:'save', minWidth: 50, maxWidth:100,isResizable: true}
      //{ key: 'column9', name: 'Scenario.', fieldName:'comments', minWidth: 50, maxWidth:100, isResizable: true, isMultiline: true}
    ];

    let migratedCaseColumns : IColumn[] = [
      { key: 'column1', name: 'Case ID', fieldName: 'caseId', minWidth:50, maxWidth:50,isResizable:true},
      { key: 'column2', name: 'Policy Number', fieldName: 'policyNumber', minWidth: 50,maxWidth:100, isResizable:true,isMultiline:true},
      { key: 'column3', name: 'Date of GK Note', fieldName: 'dueDate', minWidth: 100, maxWidth: 100,isResizable: true, isMultiline:true,
      isSorted:true, isSortedDescending:true},
      { key: 'column4', name: 'Case Status', fieldName: 'status', minWidth: 50, maxWidth: 100,isResizable: true, isMultiline:true},
      { key: 'column5', name: 'Close Case', fieldName:'closeCase', minWidth: 50, maxWidth:100, isResizable: true},
      { key: 'column6', name: 'Case Offline Files Deleted', fieldName:'delete', minWidth: 100, maxWidth:100,isResizable: true},
      { key: 'column7', name: 'Retain', fieldName:'retain', minWidth: 50, maxWidth:50,isResizable: true},
      { key: 'column8', name: '', fieldName:'save', minWidth: 50, maxWidth:100,isResizable: true}
    ];

    caseRecords.push(
      {Title:'mig-Case:6237890',CaseModifiedDate : modifiedDate,ID :3674,Retained_Or_Deleted : 'Delete',URN : 'URN:65479',
      DeleteConsentDate : new Date(2020,12,30),DeleteConsentBy : 'Rohit V'},
      {Title:'mig-Case:628610',CaseModifiedDate : modifiedDate,ID :3813,Retained_Or_Deleted : 'Retain',
      RetainConsentDate : new Date(2020,11,30),RetainConsentBy : 'Rohit V', NextReviewDate : new Date(2021, 3, 14)},
      {Title:'Case:13824',CaseModifiedDate : modifiedDate,ID :13824,Retained_Or_Deleted : 'Retain',
      RetainConsentDate : new Date(2020,11,30),RetainConsentBy : 'Rohit V'},
      {Title:'Case:13890',CaseModifiedDate : modifiedDate,ID :13890,Retained_Or_Deleted : 'Retain',
      RetainConsentDate : new Date(2020,11,30),RetainConsentBy : 'Rohit V', NextReviewDate : new Date(2021,2,25)}
      
      );

      
const currentUserPromise:Promise<ISiteUserInfo> = Promise.resolve({
  Expiration: "None",
  IsEmailAuthenticationGuestUser :true,
  UserId : {NameId :'Rohit V', NameIdIssuer: 'Unknown'},
  UserPrincipalName : null,
  Email : 'dvvvh@testdlg.com',
  Id : 5678,
  IsHiddenInUI : false,
  IsShareByEmailGuestUser : false,
  IsSiteAdmin : false,
  LoginName : 'Rohit V',
  PrincipalType : 2,
  Title : 'Rohit.V'
});
jest.mock('@pnp/sp-commonjs', () => ({
sp: {
  web: {
    currentUser: {
      get: () => currentUserPromise
    },
    lists: {
      getByTitle: (title) => ({
        items: {
          getAll: () => caseRecords,
          select: (fields) => ({
            orderBy: (fieldName, ascending) => ({
              top: (count) => ({
                  get: () => [{ID:5078}]
              })
            }),
            
            
          })
          
        },
        renderListDataAsStream: (parameters) => ({Row:[{
                  ID:5831,DeletedDate: new Date(2020,11,24),CaseID:'14,623'}]})
      })
    }
    
    
      
  },
},
})); 



configure({ adapter: new Adapter() });


describe("Enzyme basics", () => {
  let reactComponent: ShallowWrapper<IPivotTabsLargeExampleProps, IPivotTabsLargeExampleState>;
  
  
  beforeEach(() => {
      reactComponent = shallow(React.createElement(PivotTabsLargeExample, {sp : sp}));
      jest.resetModules();
      
    });

    
  
  it('must have initial state',() => {
    const state = reactComponent.state();
    
    expect(state.normalCaseItems).toHaveLength(0);
    expect(state.retainedCaseItems).toHaveLength(3);
    expect(state.deletedCaseItems).toHaveLength(1);

    expect(state.normalCaseColumns).toHaveLength(8);
    expect(state.migratedCaseColumns).toHaveLength(8);
    expect(state.retainedCaseColumns).toHaveLength(10);
    expect(state.deletedCaseColumns).toHaveLength(5);
    //expect(state.migratedCaseItems).toHaveLength(0);
  });
  
  it('should root web part element exists', () => {
    let rootComponent = shallow(React.createElement(PivotTabsLargeExample));
    let cssSelector: string = '#parent';
    const element = rootComponent.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  

  it('should call componentDidMount', async () => {
    //const pnpjsSpy1 = jest.spyOn(sp.web.lists.getByTitle.prototype.items,"getAll");
    //const pnpjsSpy1 = jest.spyOn(sp.web.lists,"getByTitle");
    //pnpjsSpy1.mockImplementation((title) => {})
    //const pnpjsSpy = jest.spyOn(sp.web.currentUser, "get");
    const instance = reactComponent.instance() as PivotTabsLargeExample;
    instance.componentDidMount();
  });
  

  it('works with async/await and resolves', async () => {
    //expect.assertions(1);
    reactComponent.update();
    
    let normalCaseRecords : any[] = [];
    const instance = reactComponent.instance() as PivotTabsLargeExample;
    const getNormalCaseItemSpy = jest.spyOn(instance, 'getNormalCaseItemsForRetention');
    getNormalCaseItemSpy.mockImplementation();
    setTimeout(() => {
      expect(getNormalCaseItemSpy).toHaveBeenCalled();     
      },10);
      
      getNormalCaseItemSpy.mockRestore();  
  });

  it('should call getNormalCasesForRetention', async () => {
    //expect.assertions(1);
    reactComponent.update();
                                     
    let normalCaseRecords : any[] = [];
    normalCaseRecords.push(
      {CaseModifiedDate : modifiedDate,ID :13844,Retained_Or_Deleted : 'Delete',URN : 'URN:65479',
      DeleteConsentDate : new Date(2020,12,30),DeleteConsentBy : 'Rohit V'},
      {CaseModifiedDate : modifiedDate,ID :13824,Retained_Or_Deleted : 'Retain',
      RetainConsentDate : new Date(2020,11,30),RetainConsentBy : 'Rohit V'}
      
      );
    
    const instance = reactComponent.instance() as PivotTabsLargeExample;
    instance.getNormalCaseItemsForRetention(normalCaseRecords, thresholdDate); 
  });

  /* it('should call getDeletedCaseItems', async () => {
    const pnpjsSpy1 = jest.spyOn(sp.web.lists,"getByTitle");
    /* pnpjsSpy1.mockImplementationOnce((title) => ({
        items: {
          getAll: () => caseRecords,
          select: (fields) => ({
            orderBy: (fieldName, ascending) => ({
              top: (count) => ({
                  get: () => [{ID:5078}]
              })
            }),
            
            
          })
          
        },
        renderListDataAsStream: (parameters) => ({Row:[{
                  ID:5831,DeletedDate: new Date(2020,11,24),CaseID:'14,623'}]})
      })); 
    const instance = reactComponent.instance() as PivotTabsLargeExample;
    instance.getDeletedCaseItems();

  });*/
  it('should root web part element exists', () => {
    let caseComponent = shallow(React.createElement(PivotTabsLargeExample));
    
    // define the css selector  
    let cssSelector: string = '#parent';
    const element = caseComponent.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  it('should searchbox element exists', () => {
    let caseComponent = shallow(React.createElement(PivotTabsLargeExample));
    // define the css selector  
    let cssSelector: string = '#searchbox';
    const element = caseComponent.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });
  
  it('should pivot element exists', () => {
    let caseComponent = shallow(React.createElement(PivotTabsLargeExample));
    // define the css selector  
    let cssSelector: string = '#pivotelement';
    const element = caseComponent.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  it('should normal case pivot item exists', () => {
    let component = shallow(React.createElement(PivotTabsLargeExample));
    let cssSelector: string = "#normalCasePivotItem";
    const element = component.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  it('should migrated case pivot item exists', () => {
    let component = shallow(React.createElement(PivotTabsLargeExample));
    let cssSelector: string = "#migratedCasePivotItem";
    const element = component.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  it('should retained case pivot item exists', () => {
    let component = shallow(React.createElement(PivotTabsLargeExample));
    let cssSelector: string = "#retainedCasePivotItem";
    const element = component.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  it('should deleted case pivot item exists', () => {
    let component = shallow(React.createElement(PivotTabsLargeExample));
    let cssSelector: string = "#deletedCasePivotItem";
    const element = component.find(cssSelector);
    expect(element.length).toBeGreaterThan(0);
  });

  it('should call search', () => {
    const instance = reactComponent.instance() as PivotTabsLargeExample;
    //Normal Case Search
    instance._normalCaseItems = [];
    instance._normalCaseItems.push(normalCaseItem);
    instance._selectedTabItem = 'Normal Cases';
    instance.setState({itemsModified : true, normalCaseItems : instance._normalCaseItems});
    instance._onChange(null, "13831");
    instance._onChange(null, "35970");

    //Normal Case Search in else block
    instance._normalCaseItems = [];
    instance._normalCaseItems.push(normalCaseItem);
    instance._selectedTabItem = '';
    instance.setState({itemsModified : true, normalCaseItems : instance._normalCaseItems});
    instance._onChange(null, "13831");
    instance._onChange(null, "35970");

    //Migrated Case Search
    instance._migratedCaseItems = [];
    instance._migratedCaseItems.push(migratedCaseItem);
    instance._selectedTabItem = 'Migrated Cases';
    instance.setState({itemsModified : true, migratedCaseItems : instance._migratedCaseItems});
    instance._onChange(null, "3783");
    instance._onChange(null, "86239");

    //Retained Case Search
    instance._retainedCaseItems = [];
    instance._retainedCaseItems.push(migratedCaseItem);
    instance._selectedTabItem = 'Retained Cases';
    instance.setState({itemsModified : true, retainedCaseItems : instance._retainedCaseItems});
    instance._onChange(null, "3783");
    instance._onChange(null, "86239");

    //Deleted Case Search
    instance._deletedCaseItems = [];
    instance._deletedCaseItems.push(deletedCaseItem);
    instance._selectedTabItem = 'Deleted Cases';
    instance.setState({itemsModified : true, deletedCaseItems : instance._deletedCaseItems});
    instance._onChange(null, "4678");
});


});