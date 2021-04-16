import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import { WebPartContext } from '@microsoft/sp-webpart-base';  
import { ISPSearchResult } from '../components/ISearchResult';  
import { ISearchResults, ICells, ICellValue } from '../components/sharepointRespnse';  
import { escape } from '@microsoft/sp-lodash-subset';  
  
export class searchservice {    
    public getSearchResults(context: WebPartContext,query: string): Promise<ISPSearchResult[]> {  
          
        let url: string = context.pageContext.web.absoluteUrl + "/_api/search/query?querytext='" + query + "'";  
          
        return new Promise<ISPSearchResult[]>((resolve, reject) => {  
 
            this._getSearchData(context,url).then((res: ISearchResults) => {  
                let searchResp: ISPSearchResult[] = [];  
  
                // Check if there was an error  
                if (typeof res["odata.error"] !== "undefined") {  
                    if (typeof res["odata.error"]["message"] !== "undefined") {  
                        Promise.reject(res["odata.error"]["message"].value);  
                        return;  
                    }  
                }  
  
                if (!this._isNull(res)) {  
                    const fields: string = "Title,Path,Description";  
                    if (typeof res.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {  
                        if (typeof res.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {  
                            searchResp = this._setSearchResults(res.PrimaryQueryResult.RelevantResults.Table.Rows, fields);  
                        }  
                    }  
                }  
  
                // Return the retrieved result set 
                console.log(searchResp) 
                resolve(searchResp);  
            });  
        });  
    }  

    private _getSearchData(context: WebPartContext,url: string): Promise<ISearchResults> {  
        return context.spHttpClient.get(url, SPHttpClient.configurations.v1, {  
            headers: {  
                'odata-version': '3.0'  
            }  
        }).then((res: SPHttpClientResponse) => {  
            return res.json();  
        }).catch(error => {  
            return Promise.reject(JSON.stringify(error));  
        });  
    }  
 
    private _setSearchResults(crntResults: ICells[], fields: string): any[] {  
        const temp: any[] = [];  
  
        if (crntResults.length > 0) {  
            const flds: string[] = fields.toLowerCase().split(',');  
  
            crntResults.forEach((result) => {  
                // Create a temp value  
                var val: Object = {}  
  
                result.Cells.forEach((cell:ICellValue) => {  
                    if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {  
                        // Add key and value to temp value  
                        val[cell.Key] = cell.Value;  
                    }  
                });  
  
                // Push this to the temp array  
                temp.push(val);  
            });  
        }  
  console.log(temp)
        return temp;  
    }  
    private _isNull(value: any): boolean {  
        return value === null || typeof value === "undefined";  
    } 

}  