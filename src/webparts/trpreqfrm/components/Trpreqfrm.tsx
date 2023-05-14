import * as React from 'react';

import { ITrpreqfrmProps } from './ITrpreqfrmProps';
import { ITrpreqfrmState } from './ITrpreqfrmState';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 

import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { getSP } from '../../../pnpjsconfig';
import { SPFI} from '@pnp/sp';
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { DateConvention, DateTimePicker } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import * as moment from 'moment';


//import { ProgressIndicator, Stack } from 'office-ui-fabric-react';
import "@pnp/sp/site-users/web";
import "@pnp/sp/items";
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react';
import { MessageBar, MessageBarType, Stack } from 'office-ui-fabric-react';
import { ListItemPicker} from '@pnp/spfx-controls-react';



/* const options: IDropdownOption[] = [
  
  { key: 'apple', text: 'Apple' },
  { key: 'banana', text: 'Banana' },
 
  { key: 'grape', text: 'Grape' },
 
]; */

export default class Trpreqfrm extends React.Component<ITrpreqfrmProps, ITrpreqfrmState> {
 


  constructor(props: ITrpreqfrmProps, state: ITrpreqfrmState) {  
    super(props);  
    
    this.state = {  
      title: 'Name',  
      users: [], 
      partyusers: [],
      ApplicantId:0,
      ValueDropdown:"",
      customerlist:"",
      startDate:new Date(),
      endDate:new Date(),
      dateduration:"0 Days",
      cargodescription:"",
      contractval:"",
      iscontractvalValid: true,
      portpairs:"",
      freight:"",
      othercon:"",
      applaw:"",
     showProgress:false,
      progressLabel: "File upload progress",
      progressDescription: "",
      progressPercent: 0,
      voyage:"",
      background:"",
      addothers:"",
      InterestedPartiesId:0,
      isSuccess: false,
      
      
     
     
    }; 
    
  }

  public componentDidMount()
{
  
  let email=this.props.userDisplayName;
  const _sp :SPFI = getSP(this.props.context ) ;
(_sp.web.siteUsers.getByEmail(email)()).then(user=> {this.setState({ApplicantId:user.Id})});
 

/* (_sp.web.lists.getByTitle("Transport Contract Request").items.select("ID").top(1).orderBy("ID", false)()).then((latestItemId) => {
  console.log(`Latest item ID is: ${latestItemId}`)}); */


  async function getLatestItemId() {
    const items = await _sp.web.lists.getByTitle("Transport Contract Request").items.orderBy("ID", false).top(1)();
    return items.length > 0 ? items[0].ID : null;
  }
  
  getLatestItemId().then((latestItemId) => {
    let lastitemid = "TRC-"+(latestItemId +1).toString();
    this.setState({title:lastitemid})
    console.log('Latest item ID is:',lastitemid);
  }).catch((error) => {
    console.log(`Error getting latest item ID: ${error}`);
  });



}
  public _getPeoplePickerItems=(items: any[]) =>{  
  
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
  public _getPartiesPeoplePickerItems=(items: any[]) =>{  
   console.log(items)
    let userid =items[0].id
      this.setState({ InterestedPartiesId: userid });
      console.log('Items new:', userid ); 
      /* let getSelectedUsers = [];  
      for (let item in items) {  
        getSelectedUsers.push(items[item].id);  
      }  
      this.setState({ users: getSelectedUsers });  */
     /* let selectedUsers: any[] = [];
      items.map((item) => {
        selectedUsers.push(item.id);
      });
       this.setState({users: selectedUsers});
      console.log('users:',selectedUsers)  */
      
    } 
 /*  public onDropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ValueDropdown: item.key as string});
  } */
  private _oncustomerSelectedItem=(data: { key: string; name: string }[])=> {
    
    if(data.length>0){
    this.setState({customerlist:data[0].name as string})
    }else{
      this.setState({customerlist:"No Country Selected"})
    }
    console.log("mydata",data);
    /* let getCountry = [];  

    for (let item in data) {
      getCountry.push(data[item].name); 
    }
    let strcountry:string = getCountry.toString();
    this.setState({customerlist:strcountry})
    console.log(strcountry) */
  }


  private _onofficeSelectedItem=(data: { key: string; name: string }[])=> {
    
    if(data.length>0){
    this.setState({ValueDropdown:data[0].name as string})
    }else{
      this.setState({ValueDropdown:"No Office Selected"})
    }
    console.log("mydata",data);
    /* let getCountry = [];  

    for (let item in data) {
      getCountry.push(data[item].name); 
    }
    let strcountry:string = getCountry.toString();
    this.setState({customerlist:strcountry})
    console.log(strcountry) */
  }
  private _onchangedStartDate=(stdate: any): void =>{  
    this.setState({ startDate: stdate }); 

    
    const startDate = moment(stdate);
    const timeEnd = moment(this.state.endDate);
    const diff = timeEnd.diff(startDate,'days').toString();
    //const diffDuration = moment.duration(diff)
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
   
  }
  private _onchangedEndDate=(eddate: any): void=> {  
    this.setState({ endDate: eddate });  
    const startDate = moment(this.state.startDate);
    const timeEnd = moment(eddate);
    const diff = (timeEnd.diff(startDate,'days')+1).toString();
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
  private onportpairsTextChange = (newText: string) => {
    this.setState({portpairs:newText});
   
    return newText;
 
  }
  private ontherconTextChange = (newText: string) => {
    this.setState({othercon:newText});
   
    return newText;
 
  }
  private onapplawTextChange = (newText: string) => {
    this.setState({applaw:newText});
   
    return newText;
 
  }
  private onBackgroundTextChange = (newText: string) => {
    this.setState({background:newText});
   
    return newText;
 
  }
  private onvoyageTextChange = (newText: string) => {
    this.setState({voyage:newText});
   
    return newText;
 
  }
  private onaddothersTextChange = (newText: string) => {
    this.setState({addothers:newText});
   
    return newText;
 
  }
  private _onccontractval=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
    const isNumberValid: boolean = !isNaN(Number(newText));
    this.setState({contractval:newText || '', iscontractvalValid: isNumberValid})
 }

 private _onfreight=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
  this.setState({freight:newText})
}

  

  /*private uploadFile = () => {
    console.log("files")
      const _sp :SPFI = getSP(this.props.context ) ;
   let input = document.getElementById("fileInput") as HTMLInputElement;
    let file = input.files[0];
    let chunkSize = 40960; // Default chunksize is 10485760. This number was chosen to demonstrate file upload progress
    this.setState({ showProgress: true });
    _sp.web.getFolderByServerRelativePath("Shared Documents").files.addChunked(file.name, file,
        data => {
          let percent = (data.blockNumber / data.totalBlocks);
          this.setState({
            progressPercent: percent,
            progressDescription: `${Math.round(percent * 100)} %`
          });
        }, true,
        chunkSize)
      .then(r => {
        console.log("File uploaded successfully");
        this.setState({
          progressPercent: 1,
          progressDescription: `File upload complete`
        });
      })
      .catch(e => {
        console.log("Error while uploading file");
        console.log(e);
      });
 
  }*/
  

    private _createItem  =async (props:ITrpreqfrmProps):Promise<void>=>{
      if (!this.state.iscontractvalValid) {
        return;
      }
    //console.log(this.props.context)
    const _sp :SPFI = getSP(this.props.context ) ;
  
      const iar =_sp.web.lists.getByTitle('Transport Contract Request').items.add({  
        
        Title: this.state.title,  
        ApplicantId: this.state.ApplicantId,
        RequestingOffice:this.state.ValueDropdown,
        Customer:this.state.customerlist,
        ContractPeriodStart:this.state.startDate,
        ContractPeriodEnd:this.state.endDate,
        ContractDuration:this.state.dateduration,
        CargoDescription:this.state.cargodescription,
        ContractVolumePerYear:this.state.contractval,
        PortPairsEstVolFreightRate:this.state.portpairs,
        FreightPayment:this.state.freight,
        OtherConditions:this.state.othercon,
        ApplicableLaw:this.state.applaw,
        VoyagePLContribution:this.state.voyage,
        Background:this.state.background,
        Others:this.state.addothers,
        InterestedPartiesId:this.state.InterestedPartiesId
      });  
      console.log('cargo added',this.state.cargodescription); 
      console.log('Item added',iar); 
   
      let input = document.getElementById("fileInput") as HTMLInputElement;
      console.log(input.files)
      //let file = input.files[0];
      let files = input.files;
      for(var i=0;i<files.length;i++)
      {

        let file = input.files[i]
        let chunkSize = 40960; // Default chunksize is 10485760. This number was chosen to demonstrate file upload progress
      this.setState({ showProgress: true });
      _sp.web.getFolderByServerRelativePath("Shared Documents").files.addChunked(file.name, file,
          data => {
            let percent = (data.blockNumber / data.totalBlocks);
            this.setState({
              progressPercent: percent,
              progressDescription: `${Math.round(percent * 100)} %`
            });
          }, true,
          chunkSize)
        .then(r => {
          console.log("File uploaded successfully");
          this.setState({
            progressPercent: 1,
            progressDescription: `File upload complete`
          });
        })
        .catch(e => {
          console.log("Error while uploading file");
          console.log(e);
        });
    // catch (error) { 
    //   console.log("creation failed with error") 
       
    // } 
  }
  this.setState({ isSuccess: true });
  setTimeout(() => {this.setState({  
    title: "",  
    users: [], 
    partyusers: [],
    ApplicantId:0,
    ValueDropdown:"",
    customerlist:"",
    startDate:new Date(),
    endDate:new Date(),
    dateduration:"0 Days",
    cargodescription:"",
    contractval:"",
    iscontractvalValid: true,
    portpairs:"",
    freight:"",
    othercon:"",
    applaw:"",
   showProgress:false,
    progressLabel: "File upload progress",
    progressDescription: "",
    progressPercent: 0,
    voyage:"",
    background:"",
    addothers:"",
    InterestedPartiesId:0,
    isSuccess: false
   
  }); }, 2000);
}
      
      
   
      

  public render(): React.ReactElement<ITrpreqfrmProps> {
    let curruser:any = this.props.userDisplayName;
    const successMessage: JSX.Element | null = this.state.isSuccess ?
    <MessageBar messageBarType={MessageBarType.success}>Form submitted successfully.</MessageBar>
    : null;
    
    const textFieldErrorMessage: JSX.Element | null = !this.state.iscontractvalValid ?
      <MessageBar messageBarType={MessageBarType.error}>Please enter a valid number.</MessageBar>
      : null;
    return (
    
    <section>
      <h2>Transport Request Form</h2> 
        <div>
          <h3>Outline of the Agreement</h3>
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
      <ListItemPicker listId='08e832c7-921f-4e0c-a57a-1377f87cc596'
       context={this.props.context as any}
          columnInternalName='Title'
          keyColumnInternalName='Id'
          placeholder="Select your customer"
          substringSearch={true}
          label="Customer"
          orderBy={"Id desc"}
          itemLimit={1}
          enableDefaultSuggestions={true}
          onSelectedItem={this._oncustomerSelectedItem}
          noResultsFoundText="No Country Found"
          defaultSelectedItems = {[]}
                     />
    
    <ListItemPicker listId='e530e316-4ff9-428c-ab1e-5c1b38154ddd'
       context={this.props.context as any}
          columnInternalName='Title'
          keyColumnInternalName='Id'
          placeholder="Select"
          substringSearch={true}
          label="Requesting Office"
          orderBy={"Id desc"}
          itemLimit={1}
          enableDefaultSuggestions={true}
          onSelectedItem={this._onofficeSelectedItem}
          noResultsFoundText="No office Found"
          defaultSelectedItems = {[]}
                     />
   
      <Stack horizontal>
      <DateTimePicker 
    
      label="From"
      maxDate={this.state.endDate}
            dateConvention={DateConvention.Date}
            value={this.state.startDate}  
            onChange={this._onchangedStartDate} 
          
            />
                
           
    <DateTimePicker label="To"
          dateConvention={DateConvention.Date}
          value={this.state.endDate}  
          onChange={this._onchangedEndDate}  />
    </Stack>  

      <TextField label="Contract Duration" value={this.state.dateduration} onChange={this._onchangedduration}/>

      <RichText label="Cargo Description" value={this.state.cargodescription}  onChange={(text)=>this.oncargodescTextChange(text)} />
      <TextField label="Contract Volume Per Year" value={this.state.contractval} onChange={this._onccontractval} errorMessage={textFieldErrorMessage?.props.messageBarType === MessageBarType.error ? textFieldErrorMessage.props.children : undefined} /> 
     
      <RichText label="Port Pairs, Estimate Volume & Freight Rate" value={this.state.portpairs}  onChange={(text)=>this.onportpairsTextChange(text)}/> 
    
      <TextField label="Freight Payment" value={this.state.freight} onChange={this._onfreight}/>
      <RichText label="Other Conditions" value={this.state.othercon}  onChange={(text)=>this.ontherconTextChange(text)}/> 
      <RichText label="Applicable Law" value={this.state.applaw}  onChange={(text)=>this.onapplawTextChange(text)}/> <br/>
      <PrimaryButton text="Show Additional Information" />
      </div>
      <div>
      <h3>Additional Information</h3>
      <RichText label="Background" value={this.state.background}  onChange={(text)=>this.onBackgroundTextChange(text)}/>
      <br />
      <input type="file" id="fileInput" name='files[]' multiple/><br />
      
        {/* <PrimaryButton text="Upload" onClick={this.uploadFile} /> <br />
        <ProgressIndicator
          label={this.state.progressLabel}
          description={this.state.progressDescription}
          percentComplete={this.state.progressPercent}
          barHeight={5} /> */}
    

      <RichText label="Voyage P/L Contribution" value={this.state.voyage}  onChange={(text)=>this.onvoyageTextChange(text)}/> 
      <br />
      {/* <PrimaryButton text="Upload" onClick={this.uploadFile} /> <br /> */}
      <input type="file" id="fileInput" multiple/><br />
    <RichText label="Others" value={this.state.addothers}  onChange={(text)=>this.onaddothersTextChange(text)}/> 
    <br />
    <input type="file" id="fileInput" multiple/>
   <br />
  
    {/* <PrimaryButton text="Upload" onClick={this.uploadFile} /> <br /> */}
   {/*  <UploadFiles
          pageSize={5}
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
          
        /> */}
      
    <PeoplePicker
        context={this.props.context as any}
        titleText="Interested Parties"
        personSelectionLimit={5}
        groupName={""} 
        showtooltip={false}
        ensureUser={true}
        required={false}
        disabled={false}
        onChange={this._getPartiesPeoplePickerItems}
        showHiddenInUI={false}
        principalTypes={[PrincipalType.User]}
        resolveDelay={1000} />                  
    <br/>
    
    <Stack horizontal horizontalAlign='end'>     
    <PrimaryButton text="Submit" onClick={() => this._createItem(this.props)} />
    {successMessage}
    </Stack> 
        </div>
      </section>
    );
  }
}




