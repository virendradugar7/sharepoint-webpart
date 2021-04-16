import * as React from 'react';
import styles from './Spfxw2.module.scss';
import { sp } from "@pnp/sp/presets/all"; 
import { Web } from "@pnp/sp/webs";  
import { ISPSearchResult } from '../components/ISearchResult';  
import { Link } from 'office-ui-fabric-react/lib/Link';  
import { ISpfxw2Props } from './ISpfxw2Props';
import { IButtonProps, DefaultButton } from 'office-ui-fabric-react/lib/Button';  
import { ISpfxw2State } from './ISpfxw2State';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';  
import { List } from 'office-ui-fabric-react/lib/List';
import { TextField } from 'office-ui-fabric-react/lib/TextField';  

import { escape } from '@microsoft/sp-lodash-subset';
import {operations} from "../Services/Services";
import  {searchservice} from "../Services/searchservice"
import {BaseButton, Dropdown, IDropdownOption,SearchBox} from 'office-ui-fabric-react'
export interface IDataFromOtherScState {  
  listItems: any;  
} 
export default class Spfxw2 extends React.Component<ISpfxw2Props,ISpfxw2State, {}> {
  private searchservice: searchservice;
  public op:operations;
 public selectedtitle:string;

  constructor(props:ISpfxw2Props){
      super(props);
      this.op=new operations();
      this.state={optionslist:[],searchText:"",searchresult:{},items:[]};
      this.searchservice=new searchservice();
// this.handlesearch=this.handlesearch.bind(this);
this._onRenderCell=this._onRenderCell.bind(this);
      
  }
 
  // public search=()=>{
  //   var result:any={};
  //   result=this.ss.getSearchResults(this.props.context,this.state.searchText)
  //   this.setState({items:result});
  // console.log(this.state.items);
  
  // }    

  
  public getListTitle=(event:any,data:any)=>{
this.selectedtitle=data.text;
  }
  public componentDidMount(){
    this.op.GetAllList(this.props.context).then((result:IDropdownOption[])=>{
      this.setState({optionslist:result});
   
    });
  }
  public render(): React.ReactElement<ISpfxw2Props> {
 let option:IDropdownOption[]=[]
    return (
      <div className={ styles.spfxw2 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to my web part!</span>
              <p className={ styles.subTitle }>CURD operations</p>
            </div>
            <div id="parent" className={styles.mystyles}>
              <Dropdown options={this.state.optionslist}  className={styles.dropdown} onChange={this.getListTitle} placeholder="select list"></Dropdown>
           <BaseButton className={styles.mybutton} text="create list item" onClick={()=>this.op.CreateListItem(this.props.context,this.selectedtitle)}></BaseButton>
           <BaseButton className={styles.mybutton} text="update list item" onClick={()=>this.op.UpdateListItem(this.props.context,this.selectedtitle)}></BaseButton>
           <BaseButton className={styles.mybutton} text="delete list item" onClick={()=>this.op.DeleteListItem(this.props.context,this.selectedtitle)}></BaseButton>
           {/* <BaseButton className={styles.mybutton} text="serch" onClick={()=>this.ss.getSearchResults(this.props.context,"updated")}></BaseButton> */}
           <TextField   
                  required={true}   
                  name="txtSearchText"   
                  placeholder="Search..."  
                  value={this.state.searchText}  
                  onChanged={e => this.setState({ searchText: e })}  
              />  

              <DefaultButton  
                  data-automation-id="search"  
                  target="_blank"  
                  title="Search"  
                  onClick={()=> this.searchservice.getSearchResults(this.props.context,this.state.searchText)
                    .then((searchResp: ISPSearchResult[]): void => {
                      this.setState({
                        items: searchResp
                      });
                  })}  
                  >  
                  Search  
              </DefaultButton>

              <FocusZone direction={FocusZoneDirection.vertical}>
                <div className="ms-ListGhostingExample-container" data-is-scrollable={true}>
                  <List items={this.state.items} onRenderCell={this._onRenderCell} />
                </div>
              </FocusZone>
          </div>
          </div>
        </div>
      </div>
    );
  }
public _onRenderCell(item: ISPSearchResult, index: number, isScrolling: boolean): JSX.Element {
  return (
    <div className="ms-ListGhostingExample-itemCell" data-is-focusable={true}>
      <div className="ms-ListGhostingExample-itemContent">
        <div className="ms-ListGhostingExample-itemName">
          <Link className={styles.item} href={item.Url}>{item.Title}</Link>          
        </div>
      
        <p></p>
       
      </div>
    </div>
  );
}
}
