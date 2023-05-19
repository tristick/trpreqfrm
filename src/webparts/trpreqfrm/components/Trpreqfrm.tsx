import * as React from 'react';

import { ITrpreqfrmProps } from './ITrpreqfrmProps';
import { ITrpreqfrmState } from './ITrpreqfrmState';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import styles from "./Trpreqfrm.module.scss"

//import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
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
import { IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles, Label, MessageBar, MessageBarType, Stack, getId } from 'office-ui-fabric-react';
import { ListItemPicker} from '@pnp/spfx-controls-react';
import * as ReactDOM from 'react-dom';
import "@pnp/sp/folders";
import * as formconst from "../../constant";
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css'; 





/* const options: IDropdownOption[] = [
  
  { key: 'apple', text: 'Apple' },
  { key: 'banana', text: 'Banana' },
 
  { key: 'grape', text: 'Grape' },
 
]; */

/* const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested colors',
  noResultsFoundText: 'No color tags found',
};

const testTags: ITag[] = [
  'black',
  'blue',
  'brown',
  'cyan',
  'green',
  'magenta',
  'mauve',
  'orange',
  'pink',
  'purple',
  'red',
  'rose',
  'violet',
  'white',
  'yellow',
].map((item) => ({ key: item, name: item[0].toUpperCase() + item.slice(1) }));

const listContainsTagList = (tag: ITag, tagList?: ITag[]) => {
  if (!tagList || !tagList.length || tagList.length === 0) {
    return false;
  }
  return tagList.some((compareTag) => compareTag.key === tag.key);
};

 */



const textFieldStyles: Partial<ITextFieldStyles> = {
  field: {
    width: '500px', // Adjust the desired width
  },
};




export default class Trpreqfrm extends React.Component<ITrpreqfrmProps, ITrpreqfrmState> {

  private dt: DataTransfer; 
 
  filesNamesRef: React.RefObject<HTMLSpanElement>;
  handleChange: any;
  pickerId: string;

  constructor(props: ITrpreqfrmProps, state: ITrpreqfrmState) {  
    super(props);  
    this.dt = new DataTransfer();
    this.filesNamesRef = React.createRef();
    this.pickerId = getId('inline-picker');
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
      bgdocuments:"",
      vdocuments:"",
      odocuments:"",
      interestedPartiesexternal: [],
      interestedPartiesexternalstr:"",
      newParty: "",
      selectedTags: [],
      baf:""
      
     
      
     
     
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
    let now = new Date();
    let formattedDate = now.toISOString().split("T")[0];
    let lastitemid = "TRC-"+(latestItemId +1)+"-"+formattedDate.toString();
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
    const diff = timeEnd.diff(startDate,'days').toString()+" Day(s)";
    //const diffDuration = moment.duration(diff)
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
   
  }
  private _onchangedEndDate=(eddate: any): void=> {  
    this.setState({ endDate: eddate });  
    const startDate = moment(this.state.startDate);
    const timeEnd = moment(eddate);
    const diff = (timeEnd.diff(startDate,'days')+1).toString()+" Day(s)";
    //const diffDuration = moment.duration(diff);
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
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
 private _onbaf=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 

  this.setState({baf:newText})

}
 private _onfreight=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
  this.setState({freight:newText})
}
private _newparty=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
  this.setState({newParty:newText})
}
handleAddParty = () => {
  const { newParty, interestedPartiesexternal } = this.state;
  if (newParty.trim() !== ''&& /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/.test(newParty)) {
    const updatedParties = [...interestedPartiesexternal, newParty]
    console.log(updatedParties)

    this.setState({ interestedPartiesexternal: updatedParties, newParty: '', interestedPartiesexternalstr:JSON.stringify(updatedParties)});
  }
  //console.log(interestedPartiesexternal.toString())
};


  bghandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    //const files = Array.prototype.slice.call(e.target.files || []);
    //console.log(files);
    //const filesNames = this.filesNamesRef.current;
  const filesNames = document.querySelector<HTMLSpanElement>('#bgfilesList > #bgfiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
          <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
    <span> x Remove </span>
  </span>
  <br/>
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
        <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
    <span> x Remove </span>
  </span>
  <br/>
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
        <span key={i} className={'file-block'}>
         <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
    <span> x Remove </span>
  </span>
  <br/>
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
     
      const upload = async () => {
        // bgfiles
        let bgfileurl = [];
        const bgcategory = 'Background'
        let bginput = document.getElementById("bgattachment") as HTMLInputElement;
    
        console.log(bginput.files);
      
        if (bginput.files.length > 0) {
          let bgfiles = bginput.files;
        
          for (var i = 0; i < bgfiles.length; i++) {
            let bgfile = bginput.files[i];
            console.log("bgfile",bgfile)
            bgfileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" +bgfile.name);
            try {
              let bguploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(bgfile.name, bgfile, (data) => {
                console.log("File uploaded successfully");
              });
              let item = await bguploadedFile.file.getItem();
              item.update({Category:bgcategory})


            } catch (err) {
              console.error("Error uploading file:", err);
            }
          }
          let convertedStr = bgfileurl.map(url => `<a href="${url.trim()}">${url.trim()}</a>`);
        let strbgurl = convertedStr.toString();
        console.log(strbgurl);
        this.setState({ bgdocuments: strbgurl });
        }
          
         else {
          console.log("No file selected for upload.");
        }
          
       
        
      
        // vfiles
        let vfileurl = [];
        let vinput = document.getElementById("vattachment") as HTMLInputElement;
        const vcategory = 'Voyage P/L Contribution'
        console.log(vinput.files);
        if (vinput.files.length > 0) {
          let vfiles = vinput.files;
        
          for (var i = 0; i < vfiles.length; i++) {
            let vfile = vinput.files[i];
            console.log("vfile",vfile)
            vfileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" + vfile.name);
            try {
              let vuploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(vfile.name, vfile, (data) => {
                console.log("File uploaded successfully");
              });
              let item = await vuploadedFile.file.getItem();
              item.update({Category:vcategory})
            } catch (err) {
              console.error("Error uploading file:", err);
            }
          }
          let vconvertedStr = vfileurl.map(url => `<a href="${url.trim()}">${url.trim()}</a>`);
        let vstrbgurl = vconvertedStr.toString();
        console.log(vstrbgurl);
        this.setState({ vdocuments: vstrbgurl });
        
        } else {
          console.log("No file selected for upload.");
          
        }
        
      
        // ofiles
        let ofileurl = [];
        let oinput = document.getElementById("othersattachment") as HTMLInputElement;
        const ocategory = 'Others'
        console.log(oinput.files);
       
        if (oinput.files.length > 0) {
          let ofiles = oinput.files;
       
          for (var i = 0; i < ofiles.length; i++) {
            let ofile = oinput.files[i];
            console.log("ofile",ofile)
            ofileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" + ofile.name);
            try {
              let ouploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(ofile.name, ofile, (data) => {
                console.log("File uploaded successfully");
              });
              let item = await ouploadedFile.file.getItem();
              item.update({Category:ocategory})      
              
            } catch (err) {
              console.error("Error uploading file:", err);
            }
          }
          let oconvertedStr = vfileurl.map(url => `<a href="${url.trim()}">${url.trim()}</a>`);
          let ostrbgurl = oconvertedStr.toString();
          console.log(ostrbgurl);
          this.setState({ odocuments: ostrbgurl });
          
        } else {
          console.log("No file selected for upload.");
          
        }
       
      }
      
      
        try {

       
         
          await upload(); // Wait for the upload function to finish
          _sp.web.lists.getByTitle(formconst.LISTNAME).items.add({
            Title: this.state.title,
            ApplicantId: this.state.ApplicantId,
            RequestingOffice: this.state.ValueDropdown,
            Customer: this.state.customerlist,
            ContractPeriodStart: this.state.startDate,
            ContractPeriodEnd: this.state.endDate,
            ContractDuration: this.state.dateduration,
            CargoDescription: this.state.cargodescription,
            ContractVolumePerYear: this.state.contractval,
            PortPairsEstVolFreightRate: this.state.portpairs,
            FreightPayment: this.state.freight,
            OtherConditions: this.state.othercon,
            ApplicableLaw: this.state.applaw,
            VoyagePLContribution: this.state.voyage,
            Background: this.state.background,
            Others: this.state.addothers,
            InterestedPartiesId: this.state.InterestedPartiesId,
            BackgroundSupportingDocuments: this.state.bgdocuments,
            VoyageP_x002f_LContributionSuppo:this.state.vdocuments,
            OthersSupportingDocuments:this.state.odocuments,
            InterestedPartiesExt:this.state.interestedPartiesexternalstr,
            BAF:this.state.baf

}).then((iar)=>{ 

  console.log('Item added',iar); });

} catch (err) {
console.error("Error creating item:", err);
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
    contractval:0,
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
    interestedPartiesexternal:[],
    newParty:""
   
  }); }, 3000);
} 
/* filterSuggestedTags = (filterText: string, tagList: ITag[]) => {
  return filterText
    ? testTags.filter(
        (tag) =>
          tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 &&
          !listContainsTagList(tag, tagList)
      )
    : [];
};
 */
getTextFromItem = (item: { name: any; }) => item.name;

handleTagChange = (selectedTags: any) => {
  this.setState({ selectedTags });
};


  public render(): React.ReactElement<ITrpreqfrmProps> {
  
   
    let curruser:any = this.props.userDisplayName;
    const {interestedPartiesexternal } = this.state;
    const successMessage: JSX.Element | null = this.state.isSuccess ?
    <MessageBar messageBarType={MessageBarType.success}>Form submitted successfully.</MessageBar>
    : null;
    
    const textFieldErrorMessage: JSX.Element | null = !this.state.iscontractvalValid ?
      <MessageBar messageBarType={MessageBarType.error}>Please enter a valid number.</MessageBar>
      : null;
    return (
    
    <section>
   
   
        <div>
          <h3>Outline of the Agreement</h3>
          <div>
    </div>
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
     {/*  <label htmlFor={this.pickerId}>Choose a color</label>
        <TagPicker
          removeButtonAriaLabel="Remove"
          //selectionAriaLabel="Selected colors"
          onResolveSuggestions={this.filterSuggestedTags}
          getTextFromItem={this.getTextFromItem}
          pickerSuggestionsProps={pickerSuggestionsProps}
          itemLimit={4}
          pickerCalloutProps={{ doNotLayer: true }}
          inputProps={{ id: this.pickerId }}
          selectedItems={selectedTags}
          onChange={this.handleTagChange}
        />
      */}
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

      {/* <TextField label="Contract Duration" value={this.state.dateduration} onChange={this._onchangedduration}/> */}
      <div><p>Contract Duration</p></div>
      <Label>{this.state.dateduration}</Label>

      {/* <RichText label="Cargo Description" value={this.state.cargodescription}  onChange={(text)=>this.oncargodescTextChange(text)} isEditMode ={true}/> */}
      <div><p>Cargo Description</p></div>
      <ReactQuill theme='snow'
    
      modules={formconst.modules}    
      formats={formconst.formats}  
      value={this.state.cargodescription}  onChange={(text)=>this.oncargodescTextChange(text)}  
        
   ></ReactQuill> 
      <TextField label="Contract Volume Per Year" value={this.state.contractval} onChange={this._onccontractval} errorMessage={textFieldErrorMessage?.props.messageBarType === MessageBarType.error ? textFieldErrorMessage.props.children : undefined} /> 
     
      {/* <RichText label="Port Pairs, Estimate Volume & Freight Rate" value={this.state.portpairs}  onChange={(text)=>this.onportpairsTextChange(text)}/>  */}
      <div><p>Port Pairs, Estimate Volume & Freight Rate</p></div>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}   
      value={this.state.portpairs}  onChange={(text)=>this.onportpairsTextChange(text)} 
        
   ></ReactQuill> 
    <TextField label="BAF" value={this.state.baf} onChange={this._onbaf}/>
      <TextField label="Freight Payment" value={this.state.freight} onChange={this._onfreight}/>
{/*       <RichText label="Other Conditions" value={this.state.othercon}  onChange={(text)=>this.ontherconTextChange(text)}/>
 
 */}     
 <div><p>Other Conditions</p></div>
  <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}   
      value={this.state.othercon}  onChange={(text)=>this.ontherconTextChange(text)}
        
   ></ReactQuill> 
    
{/*       <RichText label="Applicable Law" value={this.state.applaw}  onChange={(text)=>this.onapplawTextChange(text)}/> <br/>
 */} 
 <div><p>Applicable Law</p></div>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}   
      value={this.state.applaw}  onChange={(text)=>this.onapplawTextChange(text)}
        
   ></ReactQuill> 
    
      <PrimaryButton text="Show Additional Information" />
      </div>
      <div>
      <h3>Additional Information</h3>
      {/* <RichText label="Background" value={this.state.background}  onChange={(text)=>this.onBackgroundTextChange(text)}/> */}
      <div><p>Background</p></div>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}    
      value={this.state.background}  onChange={(text)=>this.onBackgroundTextChange(text)} 
   ></ReactQuill> 
      <br />
       
      <div id = "background" className="mt-5 text-center">
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
      
      {/* <RichText label="Voyage P/L Contribution" value={this.state.voyage}  onChange={(text)=>this.onvoyageTextChange(text)}/>  */}
      <div><p>Voyage P/L Contribution</p></div>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}  
      value={this.state.voyage}  onChange={(text)=>this.onvoyageTextChange(text)}  
        
   ></ReactQuill> 
      <br />
   
       
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
    {/* <RichText label="Others" value={this.state.addothers}  onChange={(text)=>this.onaddothersTextChange(text)}/>  */}
    <div><p>Others</p></div>
    <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}    
      value={this.state.addothers}  onChange={(text)=>this.onaddothersTextChange(text)}
        
   ></ReactQuill> 
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
 
      
    <PeoplePicker
        context={this.props.context as any}
        titleText="Interested Parties (Internal)"
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
    <Stack horizontal verticalAlign="end" className={styles.extpartiesstackContainer }>
          <TextField
            label="Interested Parties (External)"
            value={this.state.newParty}
            styles={textFieldStyles as IStyleFunctionOrObject<ITextFieldStyleProps, ITextFieldStyles>}
            onChange={this._newparty}
            onGetErrorMessage={(value) => {
              if (value && !/^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/.test(value)) {
                return 'Please enter a valid email address';
              }
              return '';
            }}
          />
          <PrimaryButton text="Add" onClick={this.handleAddParty} />
        </Stack>
        <div>
          {interestedPartiesexternal.map((party: any, index: React.Key) => (
            <span key={index}>{party}{index !== interestedPartiesexternal.length - 1 && '; '}</span>
          ))}
        </div>
        <br/>
    <Stack horizontal horizontalAlign='end' className={styles.stackContainer}>     
    <PrimaryButton text="Submit" onClick={() => this._createItem(this.props)} />
    <PrimaryButton text="Cancel"  />
    {successMessage}
    </Stack> 
        </div>
      </section>
    );
  }
}




