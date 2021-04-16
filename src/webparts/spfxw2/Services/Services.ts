import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";
// import{ISearchResult} from '../components/ISearchResult'
// import {ISearchResults} from '../components/sharepointRespnse'
export class operations {

    public GetAllList(context: WebPartContext): Promise<IDropdownOption[]> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
        var optionslist: IDropdownOption[] = [];
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                response.json().then((result: any) => {
                    console.log(result)
                    result.value.map((res: any) => {
                        console.log(res);
                        optionslist.push({ key: res.Title, text: res.Title })
                    })
                });
                resolve(optionslist);

            }, (error: any): void => {
                reject("error occusred" + error);
            })
        });


    }


    public CreateListItem(context: WebPartContext, listoption: string): Promise<string> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/items";

        const body: string = JSON.stringify({ Title: "New Title Created" });
        const options: ISPHttpClientOptions =
            { headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" }, body: body }
        return new Promise<string>
            (async (resolve, reject) => {
                context.spHttpClient.post
                (restApiUrl, SPHttpClient.configurations.v1, options);

            });

    }
    public DeleteListItem(context: WebPartContext, listoption: string): Promise<string> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/items";

        // const body: string = JSON.stringify({ Title: "New Title Created" });
        // const options: ISPHttpClientOptions =
        //     { headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" }, body: body }
        return new Promise<string>
            (async (resolve, reject) => {
                this.getLatestItemId(context,listoption).then((itemId:number)=>{
                    context.spHttpClient.post(restApiUrl+"("+itemId+")",SPHttpClient.configurations.v1,{headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" ,"IF-MATCH":"*","X-HTTP-METHOD":"DELETE"},})
                .then((response:SPHttpClientResponse)=>{
                    resolve("item with id"+itemId+"s");
                },(error:any)=>{reject("error occured");});
                });
                // context.spHttpClient.post
                // (restApiUrl, SPHttpClient.configurations.v1, options);

            });

        }
    public getLatestItemId(context:WebPartContext,listoption:string):Promise<number>{
        let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/items/?$orderby=Id desc&$top=1&select=id";
   return new Promise<number>(async(resolve,reject)=>{
context.spHttpClient.get(restApiUrl,SPHttpClient.configurations.v1,{headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "", },})
 .then((response:SPHttpClientResponse)=>{response.json().then((result:any)=>{resolve(result.value[0].Id);});}); 

    }
   )}
   public UpdateListItem(context: WebPartContext, listoption: string): Promise<string> {
    let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/item('36')";
const body:string=JSON.stringify({Title:"updated item"})
    // const body: string = JSON.stringify({ Title: "New Title Created" });
    // const options: ISPHttpClientOptions =
    //     { headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" }, body: body }
    return new Promise<string>
        (async (resolve, reject) => {

            context.spHttpClient.post(restApiUrl,SPHttpClient.configurations.v1,{headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" ,"IF-MATCH":"*","X-HTTP-METHOD":"MERGE"},body:body,})
            .then((response:SPHttpClientResponse)=>{
                resolve("item updated");
            },(error:any)=>{reject("error occured");});

        });

    }
//     public searchget(context: WebPartContext, listoption: string,searchquery: string): Promise<ISearchResult[]>{
//         // console.log(context.pageContext.web.absoluteUrl +"/_api/search/query?querytext='"+searchquery+"'&selectedProperties='Title''")
       
//         // let restApiUrl: string = context.pageContext.web.absoluteUrl +"/_api/search/query?querytext='"+searchquery+"'&selectedProperties='Title''",enablefql=true;
//         let restApiUrl: string = context.pageContext.web.absoluteUrl +"/_api/search/query"
//     return context.spHttpClient.get(restApiUrl,SPHttpClient.configurations.v1,{
//         headers:{
//             'Aceept':'application/JSON;odata=nometadata',
//             'odata-version':''
//         }
//     }).then((response:SPHttpClientResponse)=>{
//  console.log(response.json())
//         return response.json();
//     })
//     .then((responseJSON:ISearchResults)=>{
//         var result:ISearchResult[]=[];
//         var responseItems=responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows;
//         for(let res of responseItems){
//             result.push({
//                 Title:res.Cells[2].Value,
//                 ModifiedBy:res.Cells[4].Value.split('|')[1],
//             });
//         }
//         console.log(result)
//         return result;
      
//     })
//     .catch((response:SPHttpClientResponse)=>{
//         return [];
//     })
    
    
//     }
}
