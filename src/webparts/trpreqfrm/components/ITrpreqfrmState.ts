

export interface ITrpreqfrmState {  
    title: string;  
    users: string []; 
    usersstr:string
    partyusers:number[];
    ApplicantId:any;
    ValueDropdown:string ;
    customerlist:string;
    startDate:Date;
    endDate:Date;
    dateduration:any;
    cargodescription:any|string;
    contractval:number|any;
    portpairs:any|string;
    freight:string;
    othercon:string|any;
    applaw:string|any;
    showProgress:boolean;
    progressLabel: string;
    progressDescription: string;
    progressPercent: any;
    voyage:string|any;
    background:string|any;
    addothers:string|any;
    InterestedPartiesId:any;
    isSuccess: boolean;
    iscontractvalValid:boolean;
    files: any[];
    bgdocuments: string;
    vdocuments: string;
    odocuments:string;
    interestedPartiesexternal: string[],
    interestedPartiesexternalstr:string,
      newParty: string
      selectedTags: any[]
      baf:string
    
      
   
    
} 