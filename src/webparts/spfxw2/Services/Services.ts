import {sp} from "@pnp/sp/presets/all"
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";
import { Web } from "@pnp/sp/webs";
export class operations {
   
    private web = Web("https://domain07.sharepoint.com/sites/appcatalog/");
    public GetAllList(context: WebPartContext): Promise<IDropdownOption[]> {
         //let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
         
         var optionslist: IDropdownOption[] = [];
        // return new Promise<IDropdownOption[]>(async (resolve, reject) => {
        //     context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        //         response.json().then((result: any) => {
        //             console.log(result)
        //             result.value.map((res: any) => {
        //                 console.log(res);
        //                 optionslist.push({ key: res.Title, text: res.Title })
        //             })
        //         });
        //         resolve(optionslist);

        //     }, (error: any): void => {
        //         reject("error occusred" + error);
        //     })
        // });
    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
        this.web.lists.select("Title")().then((results:any)=>{
            results.map((result:any)=>{
                console.log(result);
                optionslist.push({ key: result.Title, text: result.Title });
    
            });
            resolve(optionslist);

                }, (error: any): void => {
                    reject("error occusred" + error);
                });
            });

    }


    public CreateListItem( listoption: string): Promise<string> {
        //let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/items";

        const body: string = JSON.stringify({ Title: "New Title Created" });
        const options: ISPHttpClientOptions =
            { headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" }, body: body }
        return new Promise<string>
            (async (resolve, reject) => {
                sp.web.lists.getByTitle(listoption).items.add({Title:"pnpjsitem"}).then((result:any)=>{resolve("item added succesfully")})

            });

    }
    public DeleteListItem( listoption: string): Promise<string> {
   
        return new Promise<string>
        (async (resolve, reject) => {
            this.getLatestItemId(listoption).then((itemId:number)=>{

         
            sp.web.lists.getByTitle(listoption).items.getById(itemId).delete();
        })
        });
        }
    public getLatestItemId(listoption:string):Promise<number>{
        //let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/items/?$orderby=Id desc&$top=1&select=id";
   return new Promise<number>(async(resolve,reject)=>{
sp.web.lists.getByTitle(listoption).items.select("ID").orderBy("ID",false).top(1)().then((result:any)=>{
    resolve(result[0].ID);
});
    }
   )}
   public UpdateListItem( listoption: string): Promise<string> {
    //let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listoption + "')/item('36')";
const body:string=JSON.stringify({Title:"updated item"})
    // const body: string = JSON.stringify({ Title: "New Title Created" });
    // const options: ISPHttpClientOptions =
    //     { headers: { Accept: "application/json;odata=nometadata", "content-type": "application/json;odata=nometadata", "odata-version": "" }, body: body }
    return new Promise<string>
    (async (resolve, reject) => {
        this.getLatestItemId(listoption).then((itemId:number)=>{
        sp.web.lists.getByTitle(listoption).items.getById(itemId).update({Title:"updated"});
    });
    });
    }
}