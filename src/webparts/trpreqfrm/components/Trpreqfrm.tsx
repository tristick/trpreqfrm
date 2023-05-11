import * as React from 'react';

import { ITrpreqfrmProps } from './ITrpreqfrmProps';
import { ITrpreqfrmState } from './ITrpreqfrmState';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import { Dropdown, IDropdownOption, PrimaryButton, TextField } from '@fluentui/react';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { getSP } from '../../../pnpjsconfig';
import { SPFI} from '@pnp/sp';
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { DateConvention, DateTimePicker } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import * as moment from 'moment';
import {
  UploadFiles,
} from '@pnp/spfx-controls-react/lib/UploadFiles';


const options: IDropdownOption[] = [
  
  { key: 'apple', text: 'Apple' },
  { key: 'banana', text: 'Banana' },
 
  { key: 'grape', text: 'Grape' },
 
];

export default class Trpreqfrm extends React.Component<ITrpreqfrmProps, ITrpreqfrmState> {
 
  constructor(props: ITrpreqfrmProps, state: ITrpreqfrmState) {  
    super(props);  
   
    this.state = {  
      title: '123',  
      users: [], 
      ApplicantId:0,
      ValueDropdown:"",
      startDate:new Date(),
      endDate:new Date(),
      dateduration:0,
      cargodescription:"",
      contractval:""
      
     
    }; 
    
  }


  public _getPeoplePickerItems=(items: any[]) =>{  
  console.log("m here")
  let userid =items[0].id
    this.setState({ ApplicantId: userid });
    console.log('Items new:', userid );
    /* let getSelectedUsers = [];  
    for (let item in items) {  
      getSelectedUsers.push(items[item].id);  
    }  
    this.setState({ users: getSelectedUsers });  */
   /*  let selectedUsers: any[] = [];
    items.map((item) => {
      selectedUsers.push(item.id);
    });
     this.setState({users: selectedUsers});
    console.log('users:',selectedUsers)  */
    
  } 
  public onDropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ValueDropdown: item.key as string});
  }
  private _onchangedStartDate=(stdate: any): void =>{  
    this.setState({ startDate: stdate }); 

    
    const startDate = moment(stdate);
    const timeEnd = moment(this.state.endDate);
    const diff = timeEnd.diff(startDate,'days');
    //const diffDuration = moment.duration(diff)
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
   
  }
  private _onchangedEndDate=(eddate: any): void=> {  
    this.setState({ endDate: eddate });  
    const startDate = moment(this.state.startDate);
    const timeEnd = moment(eddate);
    const diff = timeEnd.diff(startDate,'days')+1;
    //const diffDuration = moment.duration(diff);
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
  }


  private _onchangedduration=(): void =>{ 
     console.log("ch")
  }
  private oncargodescTextChange = (newText: string) => {
    this.setState({cargodescription:newText});
   
    return newText;
 
  }
  private _onccontractval=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
    this.setState({contractval:newText})
 }
 
  
    private _createItem  =async (props:ITrpreqfrmProps):Promise<void>=>{
    
    //console.log(this.props.context)
    const _sp :SPFI = getSP(this.props.context ) ;
  
      const iar =_sp.web.lists.getByTitle('Transport Contract Request').items.add({  
        
        Title: this.props.userDisplayName,  
        ApplicantId: this.state.ApplicantId,
        RequestngOffice:this.state.ValueDropdown,
        ContractPeriodFrom:this.state.startDate,
        ContractPeriodTo:this.state.endDate,
        ContractDuration:this.state.dateduration,
        CargoDescription:this.state.cargodescription,
        ContractVolumePerYear:this.state.contractval
      });  
      console.log('cargo added',this.state.cargodescription); 
      console.log('Item added',iar); 
   
    
    // catch (error) { 
    //   console.log("creation failed with error") 
       
    // } 
  } 

  public render(): React.ReactElement<ITrpreqfrmProps> {
    let curruser:any = this.props.userDisplayName
    return (
    
    <section>
        <div>
        <PeoplePicker
    context={this.props.context as any}

    titleText="Applicant"
    placeholder='Select Applicant'
    defaultSelectedUsers = {[curruser]}
    personSelectionLimit={1}
    groupName={""} // Leave this blank in case you want to filter from all users
    ensureUser={true}
    showtooltip={false}
    suggestionsLimit={5}
    required={false}
    disabled={false}
    onChange={this._getPeoplePickerItems}
    showHiddenInUI={false}
    principalTypes={[PrincipalType.User]}
    resolveDelay={1000}
    />
    <Dropdown
        placeholder="Select"
        label="Requesting Office"
        options={options}
        onChange={this.onDropdownChange}
      />
      <DateTimePicker label="From"
                dateConvention={DateConvention.Date}
                value={this.state.startDate}  
                onChange={this._onchangedStartDate} 
               
                />
                
                
<DateTimePicker label="To"
                dateConvention={DateConvention.Date}
                value={this.state.endDate}  
          onChange={this._onchangedEndDate}  />

<TextField label="Contract Duration" value={this.state.dateduration} onChange={this._onchangedduration}/>

<RichText label="Cargo Description" value={this.state.cargodescription}  onChange={(text)=>this.oncargodescTextChange(text)} />
<TextField label="Contract Volume Per Year" value={this.state.contractval} onChange={this._onccontractval}/> 
<PrimaryButton onClick={() => this._createItem(this.props)} text="Submit" />
<UploadFiles
          pageSize={20}
          context={this.props.context as any}
          title="Upload Files"
          onUploadFiles={(files) => {
            console.log("files", files);
            const _sp :SPFI = getSP(this.props.context ) ;
            files.forEach(function (value) {
            //let file = files[0];
        _sp.web.getFolderByServerRelativePath("Shared Documents").files.addChunked(value.name,value as Blob , data => {
              console.log(`progress`);
              }, true);})
          }}
          
        />
        <UploadFiles
          pageSize={20}
          context={this.props.context as any}
          title="Upload Files"
          onUploadFiles={(files) => {
            console.log("files", files);
            const _sp :SPFI = getSP(this.props.context ) ;
            files.forEach(function (value) {
            //let file = files[0];
        _sp.web.getFolderByServerRelativePath("Shared Documents").files.addChunked(value.name,value as Blob , data => {
              console.log(`progress`);
              }, true);})
          }}
          
        />
        </div>
      </section>
    );
  }
}



