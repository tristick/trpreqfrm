import * as React from 'react';
//import * as styles from './Trpreqfrm.module.scss'
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
import * as ReactDOM from 'react-dom';
import "@pnp/sp/folders";
import * as formconst from "../../constant";


/* const options: IDropdownOption[] = [
  
  { key: 'apple', text: 'Apple' },
  { key: 'banana', text: 'Banana' },
 
  { key: 'grape', text: 'Grape' },
 
]; */

export default class Trpreqfrm extends React.Component<ITrpreqfrmProps, ITrpreqfrmState> {

  private dt: DataTransfer; 
 
  filesNamesRef: React.RefObject<HTMLSpanElement>;

  constructor(props: ITrpreqfrmProps, state: ITrpreqfrmState) {  
    super(props);  
    this.dt = new DataTransfer();
    this.filesNamesRef = React.createRef();
    this.state = {  
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
      isSuccess: false,
      files: [],
      bgdocuments:""
     
      
     
     
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
    const items = await _sp.web.lists.getByTitle(formconst.LISTNAME).items.orderBy("ID", false).top(1)();
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
  
  bghandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    //const files = Array.prototype.slice.call(e.target.files || []);
    //console.log(files);
    //const filesNames = this.filesNamesRef.current;
  const filesNames = document.querySelector<HTMLSpanElement>('#bgfilesList > #bgfiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
          <span className="file-delete">
            <span>x Remove </span>
          </span>
          <span className="name">{e.target.files.item(i).name}</span><br/>
        </span>
      );
      //const fileBlocNode = fileBloc as unknown as Node; // convert to Node
      //console.log(fileBlocNode);
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      //filesNames?.appendChild(fileBloc);
      }
    }
  
    for (let file of e.target.files as any) {
      this.dt.items.add(file);
    }
  
    e.target.files = this.dt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.dt.items.length; i++) {
          if (name === this.dt.items[i].getAsFile()?.name) {
            this.dt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.dt.files;
      });
    });
  };

  vhandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    //const files = Array.prototype.slice.call(e.target.files || []);
    //console.log(files);
    //const filesNames = this.filesNamesRef.current;
  const filesNames = document.querySelector<HTMLSpanElement>('#vfilesList > #vfiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
          <span className="file-delete">
            <span>x Remove </span>
          </span>
          <span className="name">{e.target.files.item(i).name}</span><br/>
        </span>
      );
      //const fileBlocNode = fileBloc as unknown as Node; // convert to Node
      //console.log(fileBlocNode);
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      //filesNames?.appendChild(fileBloc);
      }
    }
  
    for (let file of e.target.files as any) {
      this.dt.items.add(file);
    }
  
    e.target.files = this.dt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.dt.items.length; i++) {
          if (name === this.dt.items[i].getAsFile()?.name) {
            this.dt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.dt.files;
      });
    });
  };
  
  ohandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    //const files = Array.prototype.slice.call(e.target.files || []);
    //console.log(files);
    //const filesNames = this.filesNamesRef.current;
  const filesNames = document.querySelector<HTMLSpanElement>('#ofilesList > #ofiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
          <span className="file-delete">
            <span>x Remove </span>
          </span>
          <span className="name">{e.target.files.item(i).name}</span><br/>
        </span>
      );
      //const fileBlocNode = fileBloc as unknown as Node; // convert to Node
      //console.log(fileBlocNode);
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      //filesNames?.appendChild(fileBloc);
      }
    }
  
    for (let file of e.target.files as any) {
      this.dt.items.add(file);
    }
  
    e.target.files = this.dt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.dt.items.length; i++) {
          if (name === this.dt.items[i].getAsFile()?.name) {
            this.dt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.dt.files;
      });
    });
  };
    private _createItem  =async (props:ITrpreqfrmProps):Promise<void>=>{
      const _sp :SPFI = getSP(this.props.context ) ;
      let folderUrl: string;
      if (!this.state.iscontractvalValid) {
        return;
      }
    
    let folderName =this.state.title; 
      folderUrl =formconst.LIBRARYNAME + "/" + folderName    
      _sp.web.folders.addUsingPath(folderUrl);
     
     _sp.web.lists.getByTitle(formconst.LISTNAME).items.add({  
        
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
        InterestedPartiesId:this.state.InterestedPartiesId,
        //BackgroundSupportingDocuments:this.state.bgdocuments

      }).then((iar)=>{ 
      console.log('cargo added',this.state.cargodescription); 
      console.log('Item added',iar); 

      //bgfiles
      let bgfileurl=[];
      let bginput = document.getElementById("bgattachment") as HTMLInputElement;
      console.log(bginput.files)

      if (bginput.files.length === 0) {
        console.log("No file selected for upload.");
        
      }else{
       //let file = input.files[0];
      let bgfiles = bginput.files;
      for(var i=0;i<bgfiles.length;i++)
      {
        let bgfile = bginput.files[i]
        bgfileurl.push(formconst.WEB_URL+"/"+folderUrl+ bgfile.name)
        try {
          _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(bgfile.name, bgfile, data => {
            console.log("File uploaded successfully");
          });
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }}
      let strbgurl = bgfileurl.toString();
      console.log(bgfileurl)
      this.setState({ bgdocuments: strbgurl });


      //vfiles
      let vinput = document.getElementById("vattachment") as HTMLInputElement;
      console.log(vinput.files)
      if (vinput.files.length === 0) {
        console.log("No file selected for upload.");
       
      }else{
      //let file = input.files[0];
      let vfiles = vinput.files;
      for(var i=0;i<vfiles.length;i++)
      {

        let vfile = vinput.files[i]
        try {
          _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(vfile.name, vfile, data => {
            console.log("File uploaded successfully");
          });
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }}
      

       //ofiles 
        let oinput = document.getElementById("othersattachment") as HTMLInputElement;
      console.log(oinput.files)
      //let file = input.files[0];
      let ofiles = oinput.files;
      if (oinput.files.length === 0) {
        console.log("No file selected for upload.");
        
      }else{
      for(var i=0;i<ofiles.length;i++)
      {

        let ofile = oinput.files[i]
        try {
          _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(ofile.name, ofile, data => {
            console.log("File uploaded successfully");
          });
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }}

    }).catch((error) => {
      console.log(error);
  });
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
   
  }); }, 3000);
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
      <ListItemPicker listId={formconst.CUSTOMER_LIST_ID}
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
    
    <ListItemPicker listId={formconst.REPORTING_OFFICE_LIST_ID}
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
    minDate={this.state.startDate}
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
       
      <div className="mt-5 text-center">
        <label htmlFor="bgattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="bgattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.bghandleFileUpload}
        />

        <p id="bgfiles-area">
          <span id="bgfilesList">
            <span ref={this.filesNamesRef} id="bgfiles-names"></span>
          </span>
        </p>
      </div>
      
      <RichText label="Voyage P/L Contribution" value={this.state.voyage}  onChange={(text)=>this.onvoyageTextChange(text)}/> 
      <br />
      {/* <PrimaryButton text="Upload" onClick={this.uploadFile} /> <br /> */}
       
      <div className="mt-5 text-center">
        <label htmlFor="vattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="vattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.vhandleFileUpload}
        />

      <p id="vfiles-area">
          <span id="vfilesList">
            <span ref={this.filesNamesRef} id="vfiles-names"></span>
          </span>
        </p>
      </div>
    <RichText label="Others" value={this.state.addothers}  onChange={(text)=>this.onaddothersTextChange(text)}/> 
    <br />
    <div className="mt-5 text-center">
        <label htmlFor="othersattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="othersattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.ohandleFileUpload}
        />

<p id="ofiles-area">
          <span id="ofilesList">
            <span ref={this.filesNamesRef} id="ofiles-names"></span>
          </span>
        </p>
      </div>
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
    <PrimaryButton text="Cancel"  />
    {successMessage}
    </Stack> 
        </div>
      </section>
    );
  }
}




