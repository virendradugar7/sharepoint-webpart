import * as React from 'react';
import styles from './Spfxw2.module.scss';
import { ISpfxw2Props } from './ISpfxw2Props';
import { ISpfxw2State } from './ISpfxw2State';
import { escape } from '@microsoft/sp-lodash-subset';
import {operations} from "../Services/Services";
import {BaseButton, Dropdown, IDropdownOption} from 'office-ui-fabric-react'
export default class Spfxw2 extends React.Component<ISpfxw2Props,ISpfxw2State, {}> {
 public op:operations;
 public selectedtitle:string;
  constructor(props:ISpfxw2Props){
      super(props);
      this.op=new operations();
      this.state={optionslist:[]};
      
  }
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
          </div>
          </div>
        </div>
      </div>
    );
  }
}
