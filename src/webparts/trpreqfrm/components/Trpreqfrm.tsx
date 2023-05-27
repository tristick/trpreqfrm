import * as React from 'react';

import { ICustomer, ITrpreqfrmProps } from './ITrpreqfrmProps';
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
import { getCustomerItems, getCustomerRef, getOfficeRef, submitDataAndGetId, updateData } from '../../../services/formservices';





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


  let listId: number;
  let customerreference:string;
  let officereference:string;
  let isselectedApplicant:boolean = true ;
  let isselectedCustomer:boolean = true ;
  let isselectedOffice:boolean = true ;
  let isemailInvalid:boolean = false;
  let isbuttondisbled : boolean = false;
  let buttontext : string = "Submit"


export default class Trpreqfrm extends React.Component<ITrpreqfrmProps, ITrpreqfrmState> {

  private bdt: DataTransfer; 
  private vdt: DataTransfer; 
  private odt: DataTransfer; 
  
 
  filesNamesRef: React.RefObject<HTMLSpanElement>;
  handleChange: any;
  pickerId: string;
  
  constructor(props: ITrpreqfrmProps, state: ITrpreqfrmState) {  
    super(props);  
    this.bdt = new DataTransfer();
    this.vdt = new DataTransfer();
    this.odt = new DataTransfer();
    this.filesNamesRef = React.createRef();
    this.pickerId = getId('inline-picker');
    this.state = {  
      title: "",  
      users: [], 
      usersstr:"",
      partyusers: [],
      ApplicantId:0,
      //selectedApplicant: "default",
      ValueDropdown:"",
      //selectedOffice: null,
      customerlist:"",
      //selectedCustomer: null,
      //isselectedCustomer: true,
      startDate:new Date(),
      endDate:new Date(),
      dateduration:"0",
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
      files: [],
      bgdocuments:"",
      vdocuments:"",
      odocuments:"",
      interestedPartiesexternal: [],
      interestedPartiesexternalstr:"",
      newParty: "",
      selectedTags: [],
      baf:"",
      onload:true
      
     
      
     
     
    }; 
    
  }

  public componentDidMount()
{

  let email=this.props.userDisplayName;
  const _sp :SPFI = getSP(this.props.context ) ;
  (_sp.web.siteUsers.getByEmail(email)()).then(user=> {this.setState({ApplicantId:user.Id})});
 
 //this.fetchListId();
  
  this.fetchCustomerItems();
}

/* fetchListId = async () => {
  try {
    
    const latestItemId: ITrc[] | null = await getLatestItemId(this.props);

    if (latestItemId) {
      listId = latestItemId[0]?.ID;
      console.log('List ID:', listId);
    } else {
      listId = 0;
      console.log('No items found');
    }
  } catch (error) {
    console.error('Error fetching list ID:', error);
  }
};
 */
fetchCustomerItems = async () => {
  try {
    const customerItems: ICustomer[] = await getCustomerItems(this.props);
    console.log('Fetched customer items:', customerItems);
     /* //customerItems.forEach((customerItem) => {
    //const TitleValue = customerItem.Title; 
    //const ReferenceValue = customerItem.Reference; 
    //console.log('Title Value:', TitleValue);
    //console.log('Reference Value:', ReferenceValue);
    
  }); */
  } catch (error) {
    console.error('Error fetching customer items:', error);
  }
};


  public _getPeoplePickerItems=(items: any[]) =>{  
  
    if(items.length> 0){
      let userid =items[0].id
      this.setState({ ApplicantId: userid });
      isselectedApplicant  = true;
      //console.log('Items new:', userid );
    }else{
      this.setState({ ApplicantId: "" });
      isselectedApplicant  = false;

    }

    
  
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
    /* let userid =items[0].id
      this.setState({ InterestedPartiesId: userid });
      console.log('Items new:', userid );  */
      /* let getSelectedUsers = [];  
      for (let item in items) {  
        getSelectedUsers.push(items[item].id);  
      }  
      this.setState({ users: getSelectedUsers });  */
      let selectedUsers: string[] = [];
      items.map((item) => {
        selectedUsers.push(item.id);
       
      });
       this.setState({users: selectedUsers});
      console.log('users:',selectedUsers)  
      
    } 
 /*  public onDropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ValueDropdown: item.key as string});
  } */
  private _oncustomerSelectedItem=(data: { key: string; name: string }[])=> {

   if(data.length == 0){
    this.setState({customerlist:""})
  }else{
    this.setState({customerlist:data[0].name as string})
    getCustomerRef(this.props,data[0].name).then((customerRef: string)=>{

      customerreference = customerRef
      //console.log(customerRef);
      
    })
   }

     isselectedCustomer = data.length >0 ? true : false;
  
  }



  private _onofficeSelectedItem=(data: { key: string; name: string }[])=> {
    
    if(data.length == 0 ){
      this.setState({ValueDropdown:""})
    }else{
    this.setState({ValueDropdown:data[0].name as string})
    getOfficeRef(this.props,data[0].name).then((officeRef: string)=>{

      officereference = officeRef
      console.log(officereference);
      
    })
    }

   isselectedOffice = data.length > 0 ? true : false;
    

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

    this.setState({ interestedPartiesexternal: updatedParties, newParty: '', interestedPartiesexternalstr:(JSON.stringify(updatedParties)).slice(1, -1).replace(/"/g, '')});
    isemailInvalid = false;
  } else{

    isemailInvalid = true;
    this.setState({newParty:""})

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
     <button> Remove</button>
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
      this.bdt.items.add(file);
    }
  
    e.target.files = this.bdt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.bdt.items.length; i++) {
          if (name === this.bdt.items[i].getAsFile()?.name) {
            this.bdt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.bdt.files;
  
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
    <button> Remove</button>
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
      this.vdt.items.add(file);
    }
  
    e.target.files = this.vdt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.vdt.items.length; i++) {
          if (name === this.vdt.items[i].getAsFile()?.name) {
            this.vdt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.vdt.files;
    
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
          <button> Remove</button>
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
      this.odt.items.add(file);
    }
  
    e.target.files = this.odt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.odt.items.length; i++) {
          if (name === this.odt.items[i].getAsFile()?.name) {
            this.odt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.odt.files;
      });
    });
  };
    private _createItem  =async (props:ITrpreqfrmProps):Promise<void>=>{

      isbuttondisbled = true;
      buttontext = "Saving..."
     
     // const _sp :SPFI = getSP(this.props.context ) ;
      let folderUrl: string;
     
      if (!isselectedApplicant) {
        // Handle the validation error, e.g., display an error message
      
        return;
      }

      if ((this.state.customerlist).length == 0) {
        
        isselectedCustomer = false;
       this.setState({customerlist:""})
        //console.log(customerValidationMessage);
        return;
      }

      
      if ((this.state.ValueDropdown).length == 0) {
        
        isselectedOffice = false;
       this.setState({ValueDropdown:""})
        //console.log(customerValidationMessage);
        return;
      }

      if (isemailInvalid) {
        
      
        //console.log(customerValidationMessage);
        return;
      }

      
      /* const officeValidationMessage = this.validateOfficeField();
      if (!officeValidationMessage) {
        
        console.log(officeValidationMessage);
        return;
      }
  
      if (!this.state.iscontractvalValid) {
        return;
      }
 */
      

      let listFolderpath=formconst.WEB_URL+"/Lists/"+ formconst.LISTNAME+"/" +this.state.customerlist; 

     // console.log(listFolderpath);
      //_sp.web.folders.addUsingPath(folderUrl);
      
      const data = {
        Title: 'New Item creation in process',
     
      };
      submitDataAndGetId(this.props,data,listFolderpath).then(async (itemId: any) => {
        listId = itemId   
        console.log(`Item created with ID: ${itemId}`);

        //Request ID format
        let now = new Date();
        let options: Intl.DateTimeFormatOptions = {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      };
      let listIdstr
       if(listId < 1000 && listId > 100){
        listIdstr = "0"+String(listId)
      }else if(listId < 100){
        listIdstr ="00"+String(listId)
      } else if(listId < 10) {
        listIdstr ="000"+String(listId)
      }else{
        listIdstr = String(listId)
      }
      
      console.log(listIdstr)
      let formattedDate = now.toLocaleDateString('en-GB', options).replace(/\//g, '');;
      let lastitemid = (listIdstr)+"-"+customerreference+"-"+officereference +"-" +formattedDate.toString();

     
     // console.log(lastitemid)
    
      
    //folderUrl =formconst.LIBRARYNAME + "/" + lastitemid    
    folderUrl = formconst.LIBRARYNAME +"/" + this.state.customerlist + "/" + lastitemid
    this.setState({title:lastitemid})
    
        
   
  }).then(async () => {
    
    await upload()
    // Update the item
    const updatedData = {
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
      InterestedPartiesId: this.state.users,
      BackgroundSupportingDocuments: this.state.bgdocuments,
      VoyagePLContributionSupportingDo:this.state.vdocuments,
      OthersSupportingDocuments:this.state.odocuments,
      InterestedPartiesExt:this.state.interestedPartiesexternalstr,
      BAF:this.state.baf
    };
    return updateData(this.props,listId, updatedData);
  })
   .then(() => {
    //console.log('Item Updated successfully');
    // Perform any further actions if needed
    
    isbuttondisbled = false;
    buttontext = "Submit"
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
    dateduration:"0",
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
  }) 
  .catch((error: any) => {
    console.log('Error:', error);
  });

 
const upload = async () => {

    console.log(folderUrl)
    const _sp :SPFI = getSP(props.context) ;
    let strbgurl = "";
    let vstrbgurl = "";
    let ostrbgurl = "";
    _sp.web.folders.addUsingPath(folderUrl);
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
        //console.log()
        try {
          let bguploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(bgfile.name, bgfile, (data) => {
            console.log("File uploaded successfully");
          });
          let item = await bguploadedFile.file.getItem();
          await item.update({Section:bgcategory});
          await item.update({RequestID:this.state.title})

            
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let convertedStr = bgfileurl.map(url => `<a href="${url.trim()}">${url.trim()}</a>`);
       strbgurl = convertedStr.toString();
        //console.log(strbgurl);
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
          await item.update({Section:vcategory});
          await item.update({RequestID:this.state.title})
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let vconvertedStr = vfileurl.map(url => `<a href="${url.trim()}">${url.trim()}</a>`);
     vstrbgurl = vconvertedStr.toString();
    //console.log(vstrbgurl);
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
          await item.update({Section:ocategory});
          await item.update({RequestID:this.state.title}) ;
             
          
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let oconvertedStr = ofileurl.map(url => `<a href="${url.trim()}">${url.trim()}</a>`);
       ostrbgurl = oconvertedStr.toString();
      //console.log(ostrbgurl);
      this.setState({ odocuments: ostrbgurl });
      
    } else {
      console.log("No file selected for upload.");
      
    }

    
   
  }
    
    
      
  
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
    <MessageBar messageBarType={MessageBarType.success}>Request Id : {this.state.title} submitted successfully.</MessageBar>
    : null;
    
    const textFieldErrorMessage: JSX.Element | null = !this.state.iscontractvalValid ?
      <MessageBar messageBarType={MessageBarType.error}>Please enter a valid number.</MessageBar>
      : null;

      const applicantFieldErrorMessage: JSX.Element | null = !isselectedApplicant ?
      <MessageBar messageBarType={MessageBarType.error}>Applicant field is required.</MessageBar>
      : null;

      const customerFieldErrorMessage: JSX.Element | null = !isselectedCustomer ?
      <MessageBar messageBarType={MessageBarType.error}>Customer field is required.</MessageBar>
      : null;

      const OfficeFieldErrorMessage: JSX.Element | null = !isselectedOffice ?
      <MessageBar messageBarType={MessageBarType.error}>Office field is required.</MessageBar>
      : null;

      const EmailFieldErrorMessage: JSX.Element | null = isemailInvalid ?
      <MessageBar messageBarType={MessageBarType.error}>Please enter a valid email address.</MessageBar>
      : null;
    return (
    
    <section>
   
   
        <div>
          <p className={styles.heading}>Outline of the Agreement</p>
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
            required={true}
            disabled={false}
            onChange={this._getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
    /> {applicantFieldErrorMessage}
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
      <p className={styles.formlabel}>Customer<span className={styles.required}> *</span></p>
      <ListItemPicker listId={formconst.CUSTOMER_LIST_ID}
       context={this.props.context as any}
          columnInternalName='Title'
          keyColumnInternalName='Id'
          placeholder="Select your customer"
          substringSearch={true}
          //label="Customer *"
          orderBy={"Id desc"}
          
          itemLimit={1}
          enableDefaultSuggestions={true}
          onSelectedItem={this._oncustomerSelectedItem}
          noResultsFoundText="No Country Found"
          defaultSelectedItems = {[]}
                     />{customerFieldErrorMessage}
    <p className={styles.formlabel}>Requesting Office<span className={styles.required}> *</span></p>
    <ListItemPicker listId={formconst.REPORTING_OFFICE_LIST_ID}
       context={this.props.context as any}
          columnInternalName='Title'
          keyColumnInternalName='Id'
          placeholder="Select your office"
          substringSearch={true}
          //label="Requesting Office"
          orderBy={"Id desc"}
          itemLimit={1}
          enableDefaultSuggestions={true}
          onSelectedItem={this._onofficeSelectedItem}
          noResultsFoundText="No office Found"
          defaultSelectedItems = {[]}
         
                     />{OfficeFieldErrorMessage}
   
      <div>
      <DateTimePicker 
    
      label="Contract From"
      maxDate={this.state.endDate}
            dateConvention={DateConvention.Date}
            value={this.state.startDate}  
            onChange={this._onchangedStartDate} 
            allowTextInput = {false}
            showLabels = {false}
          
            />
                
           
    <DateTimePicker label="Contract To"
    minDate={this.state.startDate}
          dateConvention={DateConvention.Date}
          value={this.state.endDate}  
          onChange={this._onchangedEndDate}
          allowTextInput = {false}  
          showLabels = {false}/>
    </div>  

      {/* <TextField label="Contract Duration" value={this.state.dateduration} onChange={this._onchangedduration}/> */}
       <p className={styles.formlabel}>Contract Duration (Days)</p>
      <Label>{this.state.dateduration}</Label>

      {/* <RichText label="Cargo Description" value={this.state.cargodescription}  onChange={(text)=>this.oncargodescTextChange(text)} isEditMode ={true}/> */}
      <p className={styles.formlabel}>Cargo Description</p>
      <ReactQuill theme='snow'
    
      modules={formconst.modules}    
      formats={formconst.formats}  
      value={this.state.cargodescription}  onChange={(text)=>this.oncargodescTextChange(text)}  
        
   ></ReactQuill> 
      <TextField label="Contract Volume Per Year" value={this.state.contractval} onChange={this._onccontractval} errorMessage={textFieldErrorMessage?.props.messageBarType === MessageBarType.error ? textFieldErrorMessage.props.children : undefined} /> 
     
      {/* <RichText label="Port Pairs, Estimate Volume & Freight Rate" value={this.state.portpairs}  onChange={(text)=>this.onportpairsTextChange(text)}/>  */}
      <p className={styles.formlabel}>Port Pairs, Estimate Volume & Freight Rate</p>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}   
      value={this.state.portpairs}  onChange={(text)=>this.onportpairsTextChange(text)} 
        
   ></ReactQuill> 
    <TextField label="BAF" value={this.state.baf} onChange={this._onbaf}/>
      <TextField label="Freight Payment" value={this.state.freight} onChange={this._onfreight}/>
{/*       <RichText label="Other Conditions" value={this.state.othercon}  onChange={(text)=>this.ontherconTextChange(text)}/>
 
 */}     
 <p className={styles.formlabel}>Other Conditions</p>
  <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}   
      value={this.state.othercon}  onChange={(text)=>this.ontherconTextChange(text)}
        
   ></ReactQuill> 
    
{/*       <RichText label="Applicable Law" value={this.state.applaw}  onChange={(text)=>this.onapplawTextChange(text)}/> <br/>
 */} 
 <p className={styles.formlabel}>Applicable Law</p>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}   
      value={this.state.applaw}  onChange={(text)=>this.onapplawTextChange(text)}
        
   ></ReactQuill> 
    

      </div>
      <div>
      <br /><p className={styles.heading}>Additional Information</p> <br />
      {/* <RichText label="Background" value={this.state.background}  onChange={(text)=>this.onBackgroundTextChange(text)}/> */}
      <p className={styles.formlabel}>Background</p>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}    
      value={this.state.background}  onChange={(text)=>this.onBackgroundTextChange(text)} 
   ></ReactQuill> 
      
       
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
      <p className={styles.formlabel}>Voyage P/L Contribution</p>
      <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}  
      value={this.state.voyage}  onChange={(text)=>this.onvoyageTextChange(text)}  
        
   ></ReactQuill> 
      
   
       
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
    <p className={styles.formlabel}>Others</p>
    <ReactQuill theme='snow'
      modules={formconst.modules}    
      formats={formconst.formats}    
      value={this.state.addothers}  onChange={(text)=>this.onaddothersTextChange(text)}
        
   ></ReactQuill> 
    
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
   
 
      
    <PeoplePicker
        context={this.props.context as any}
        titleText="Interested Parties (MOLEA)"
        personSelectionLimit={10}
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
           /*  onGetErrorMessage={(value) => {
              if (value && !/^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/.test(value)) {
                return 'Please enter a valid email address';
              }
              return '';
            }} */
          />
          <PrimaryButton text="+" onClick={this.handleAddParty} />
          {/* <IconButton onClick= { this.handleAddParty } iconProps={ { iconName: 'Add' } } title='Add Party' /> */}
        </Stack>
    
        <div>
          {interestedPartiesexternal.map((party: any, index: React.Key) => (
            <span key={index}>{party}{index !== interestedPartiesexternal.length - 1 && '; '}</span>
          ))}
        </div>    
        <br/>
        {EmailFieldErrorMessage}
    <Stack horizontal horizontalAlign='end' className={styles.stackContainer}>     
    <PrimaryButton text={buttontext} onClick={() => this._createItem(this.props)} disabled= {isbuttondisbled}/>
    <PrimaryButton text="Cancel"  />
   
    </Stack> 
    
    {successMessage}
        </div>
      </section>
    );
  }
}




