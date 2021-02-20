var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'CreateListUsingSpfxWebPartStrings';
import styles from './components/CreateListUsingSpfx.module.scss';
import { Web } from 'sp-pnp-js';
var CreateListUsingSpfxWebPart = /** @class */ (function (_super) {
    __extends(CreateListUsingSpfxWebPart, _super);
    function CreateListUsingSpfxWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    CreateListUsingSpfxWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n\n    <div class=\"" + styles.createListUsingSpfx + "\">  \n    \n    <div class=\"" + styles.container + "\">  \n    \n    <div class=\"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row + "\">  \n    \n    <div class=\"ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1\">  \n    \n    <span class=\"ms-font-xl ms-fontColor-white\" style=\"font-size:28px\">Welcome to SPFx learning (create list using PnP JS library)</span>  \n    \n    <p class=\"ms-font-l ms-fontColor-white\" style=\"text-align: left\">Demo : Create SharePoint List in SPO using SPFx</p>  \n    \n    </div>  \n    \n    </div>  \n    \n    <div class=\"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row + "\">  \n    \n    <div data-role=\"main\" class=\"ui-content\">    \n    \n                 <div>    \n                  <input id=\"listTitle\"  placeholder=\"List Name\"/>  \n                  <button id=\"createNewCustomListToSPO\"  type=\"submit\" >Create List</button>   \n                 </div>    \n    \n               </div> \n    \n    <br>  \n    \n    <div id=\"ListCreationStatusInSPOnlineUsingSPFx\" />  \n    \n    </div>  \n    \n    </div>  \n    \n    </div>";
        this.AddEventListeners();
    };
    CreateListUsingSpfxWebPart.prototype.AddEventListeners = function () {
        var _this = this;
        document.getElementById('createNewCustomListToSPO').addEventListener('click', function () { return _this.CreateListInSPOUsinPnPSPFx(); });
    };
    CreateListUsingSpfxWebPart.prototype.CreateListInSPOUsinPnPSPFx = function () {
        var myWeb = new Web(this.context.pageContext.web.absoluteUrl);
        console.log("my web ::" + myWeb.toUrl.toString);
        //let mySPFxListTitle = "CustomList_using_SPFx_Framework"; 
        var mySPFxListTitle = document.getElementById('listTitle')["value"];
        var mySPFxListDescription = "Custom list created using the SPFx Framework";
        var listTemplateID = 100;
        var boolEnableCT = false;
        myWeb.lists.add(mySPFxListTitle, mySPFxListDescription, listTemplateID, boolEnableCT).then(function (splist) {
            document.getElementById("ListCreationStatusInSPOnlineUsingSPFx").innerHTML += "The SPO new list " + mySPFxListTitle + " has been created successfully using SPFx Framework.";
        });
        var list = myWeb.lists.getByTitle('CaseList').items.get().then(function (item) {
            console.log('list item ::' + item);
        });
        //const r = list.select('Id');
        //console.log(r); 
    };
    CreateListUsingSpfxWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(CreateListUsingSpfxWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    CreateListUsingSpfxWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return CreateListUsingSpfxWebPart;
}(BaseClientSideWebPart));
export default CreateListUsingSpfxWebPart;
//# sourceMappingURL=CreateListUsingSpfxWebPart.js.map