import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {ISPSearchResult,ISearchResults, ICells, ICellValue, ISearchResponse} from './ISearchResult';
import{EventValues} from './ISharePointSearchResults';
import {Utils} from './UrlUtils';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EventdisplayWebPart.module.scss';
import * as strings from 'EventdisplayWebPartStrings';
import * as $ from 'jquery';
import * as moment from 'moment';
require('./template.css');
import * as _ from "lodash";

export interface IEventdisplayWebPartProps {
  linkShowMore: string;
  topNResult:number;
}

export default class EventdisplayWebPart extends BaseClientSideWebPart<IEventdisplayWebPartProps> {

    private QryString:string;
    private BaseDomainUrl:string;

  constructor() {
  super();
  SPComponentLoader.loadCss('https://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700');
  SPComponentLoader.loadCss('https://use.fontawesome.com/releases/v5.3.1/css/all.css');

  }
  public render(): void {
    this.domElement.innerHTML = `
    <div class="container">
    <div class="wraper mt-5">
         <!-- main div copy and paste this section -->
         <div class="widget shadow">
             <!-- header text -->
             <div class="widget-header">Events</div>
             <!-- widget body -->
             <div class="widget_body" id="widget_body">
                
             </div>
             <!-- footer link -->
             <div class="event_footer text-right">
                <a href="${$.trim(this.properties.linkShowMore)==""?"#":$.trim(this.properties.linkShowMore)}" target="_blank">Show More <i class="fa fa-arrow-circle-right"></i></a>
             </div>
         </div>
         <!-- main div end -->
    </div>
</div>`;

    this.QryString="%27contenttype:%22Valo%20Calendar%20Event%22%27&trimduplicates=false&rowlimit=500&selectproperties=%27title%2cLocation%2cValoEventCategoryOWSCHCS%2cEventDateOWSDATE%2cEndDateOWSDATE%27";
    this.BaseDomainUrl = Utils.getAbsoluteDomainUrl();

    this.getSearchResults(this.QryString)
    .then((searchResp: ISPSearchResult[]): void => {
    
        let srchREsp:ISPSearchResult[]=  searchResp;
        console.log(srchREsp);
        let evtCollection = Array< EventValues>();
        $.each(srchREsp,(index,item)=>{
            let evtEntry=new EventValues();
            evtEntry.EventName=(<any>item).title;
            evtEntry.Location=(<any>item).Location;
            evtEntry.EventCategory=(<any>item).ValoEventCategoryOWSCHCS;
            evtEntry.StartDate= moment((<any>item).EventDateOWSDATE);
            evtEntry.EndDate= moment((<any>item).EndDateOWSDATE);

                if(evtEntry.StartDate>moment(new Date())){
                            evtCollection.push(evtEntry);
                }
        });

        evtCollection = _.sortBy(evtCollection, o=>{ return o.StartDate; });

let eventHtml:string="";
$("#widget_body").empty() ; 

console.log(evtCollection);
evtCollection=evtCollection.slice(0, this.properties.topNResult);
$.each(evtCollection,(index,item)=>{
eventHtml+=`<!-- single item -->
<div class="event_item clearfix">
    <div class="dateContainer">
        <span class="date">${item.StartDate.format("DD")}</span>
        <span class="month">${item.StartDate.format("MMM")}</span>
    </div>
    <div class="event_desc">
        <h3>${item.EventName}</h3>
        <div class="timelog"><i class="far fa-clock"></i> <span>${item.EndDate.local().format("HH:mm")}</span></div>
        <div class="timelog">${item.EventCategory}</div>
    </div>
</div>
<!-- single item -->`;
});


$("#widget_body").append(eventHtml) ;  
        
    });
  }

  public getSearchResults(query: string): Promise<ISPSearchResult[]> {  
          
    let url: string = this.BaseDomainUrl + "/_api/search/query?querytext=" + query ;  
      
    return new Promise<ISPSearchResult[]>((resolve, reject) => {  
        // Do an Ajax call to receive the search results  
        this._getSearchData(url).then((res: ISearchResults) => {  
            let searchResp: ISPSearchResult[] = [];  

            // Check if there was an error  
            if (typeof res["odata.error"] !== "undefined") {  
                if (typeof res["odata.error"]["message"] !== "undefined") {  
                    Promise.reject(res["odata.error"]["message"].value);  
                    return;  
                }  
            }  

            if (!this._isNull(res)) {  
                const fields: string = "Title,Location,ValoEventCategoryOWSCHCS,EventDateOWSDATE,EndDateOWSDATE";  

                // Retrieve all the table rows  
                if (typeof res.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {  
                    if (typeof res.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {  
                        searchResp = this._setSearchResults(res.PrimaryQueryResult.RelevantResults.Table.Rows, fields);  
                    }  
                }  
            }  

            // Return the retrieved result set  
            resolve(searchResp);  
        });  
    });  
}  

 /** 
 * Retrieve the results from the search API 
 * 
 * @param url 
 */  
private _getSearchData(url: string): Promise<ISearchResults> {  
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {  
        headers: {  
            'odata-version': '3.0'  
        }  
    }).then((res: SPHttpClientResponse) => {  
        return res.json();  
    }).catch(error => {  
        return Promise.reject(JSON.stringify(error));  
    });  
}  

/** 
 * Set the current set of search results 
 * 
 * @param crntResults 
 * @param fields 
 */  
private _setSearchResults(crntResults: ICells[], fields: string): any[] {  
    const temp: any[] = [];  

    if (crntResults.length > 0) {  
        const flds: string[] = fields.toLowerCase().split(',');  

        crntResults.forEach((result) => {  
            // Create a temp value  
            var val: Object = {};  

            result.Cells.forEach((cell: ICellValue) => {  
                if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {  
                    // Add key and value to temp value  
                    val[cell.Key] = cell.Value;  
                }  
            });  

            // Push this to the temp array  
            temp.push(val);  
        });  
    }  

    return temp;  
}  

/** 
 * Check if the value is null or undefined 
 * 
 * @param value 
 */  
private _isNull(value: any): boolean {  
    return value === null || typeof value === "undefined";  
}  

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void
  {
    $("#widget_body").empty() ;  
    this.context.propertyPane.refresh();

  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('linkShowMore', {
                  label: "Show More Link Url",
                  value:''
                }),
                PropertyPaneTextField('topNResult',{label:'Top N Results',value:'3'})
              ]
            }
          ]
        }
      ]
    };
  }
}
