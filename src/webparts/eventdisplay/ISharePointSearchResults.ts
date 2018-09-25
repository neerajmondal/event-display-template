import {ISPSearchResult} from './ISearchResult';  
import { Moment } from 'moment';
  
export interface ISearchResultsViewerState {  
    status: string;  
    searchText: string;  
    items: ISPSearchResult[];  
}  

export class EventValues{
    public EventName:string;
    public Location:string;
    public StartDate:Moment;
    public EndDate:Moment;
    public EventCategory:string;

}