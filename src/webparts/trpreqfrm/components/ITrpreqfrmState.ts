

export interface ITrpreqfrmState {  
    title: string;  
    users: string []; 
    //selectedApplicant: any,
    usersstr:string
    partyusers:number[];
    ApplicantId:any;
    ValueDropdown:string ;
    //selectedOffice: any,
    customerlist:string;
    //selectedCustomer: any,
    //isselectedCustomer:boolean,
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
    onload:boolean,
    allfieldsvalid:boolean,
    listfolderExists: boolean;
    libfolderExists: boolean

    
      
   
    
} 