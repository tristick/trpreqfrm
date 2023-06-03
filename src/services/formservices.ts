import { SPFI } from "@pnp/sp";

import { getSP } from "../pnpjsconfig";
import { ICustomer, ITrc, ITrpreqfrmProps } from "../webparts/trpreqfrm/components/ITrpreqfrmProps";
import * as formconst from "../webparts/constant";






export const getCustomerItems= async (props:ITrpreqfrmProps):Promise<ICustomer[]>=> {
    const _sp :SPFI = getSP(props.context) ;
    const items =_sp.web.lists.getByTitle(formconst.CUSTOMER_LISTNAME).items.select().orderBy('Title',true)();
    console.log('Customer items',items);
    return items;
      
      
  }
  
  export const getLatestItemId= async (props:ITrpreqfrmProps):Promise<ITrc[]>=> {
    const _sp :SPFI = getSP(props.context) ;
    const items = await _sp.web.lists.getByTitle(formconst.LISTNAME).items.orderBy("ID", false).top(1)();
    return items;
    //return items.length > 0 ? items[0].ID : 0;
  }

  export const getCustomerRef=(props:ITrpreqfrmProps,customerName: string) => {
    console.log(customerName)
    const _sp :SPFI = getSP(props.context) ;
    return new Promise((resolve, reject) => {
      _sp.web.lists.getByTitle(formconst.CUSTOMER_LISTNAME).items.select("RefCode").filter(`Title eq '${customerName}'`)()
        .then((items) => {
          if (items.length > 0) {
            const customerRef = items[0].RefCode;
            console.log(customerRef)
            resolve(customerRef);
          } else {
            reject(new Error("Customer not found"));
          }
        })
        .catch((error) => {
          reject(error);
        });
    });
  }

  export const submitDataAndGetId = async (props:ITrpreqfrmProps,data:any,listFolderpath: string): Promise<any> => {
  
    const _sp :SPFI = getSP(props.context) ;
    _sp.web.folders.addUsingPath(listFolderpath);
    return _sp.web.lists.getByTitle(formconst.LISTNAME).addValidateUpdateItemUsingPath([
      { FieldName: 'Title', FieldValue: data.Title },
      { FieldName: 'CargoDescription', FieldValue: data.CargoDescription },
      { FieldName: 'PortPairsEstVolFreightRate', FieldValue: data.PortPairsEstVolFreightRate },
      { FieldName: 'OtherConditions', FieldValue: data.OtherConditions },
      { FieldName: 'ApplicableLaw', FieldValue: data.ApplicableLaw },
      { FieldName: 'VoyagePLContribution', FieldValue: data.VoyagePLContribution },
      { FieldName: 'Background', FieldValue: data.Background },
      { FieldName: 'Others', FieldValue: data.Others }
    ], listFolderpath)
      .then((response) => {
        console.log(response)
        //console.log("New Item",response[8].FieldValue)
      
          // Send item ID as a response to the promise
          const itemId = response[8].FieldValue;
          console.log("ID",itemId)
          // Resolve the promise with the item ID
          return Promise.resolve(itemId);
      })
      .catch((error) => {
          // Handle any errors that occurred during the request
          return Promise.reject(error);
      });

    
  }

  export const updateData=(props:ITrpreqfrmProps ,itemId: number, data: any): Promise<void>=> {
    const _sp :SPFI = getSP(props.context) ;
    return new Promise<void>((resolve, reject) => {
      _sp.web.lists.getByTitle(formconst.LISTNAME).items.getById(itemId).update(data)
        .then(() => {
          
          //console.log(e.response.headers.get("content-length"))
          resolve();
        })
        .catch((error) => {
   
          reject(error);
        });
    });
  }


  export const getOfficeRef=(props:ITrpreqfrmProps,officeName: string) => {
    console.log(officeName)
    const _sp :SPFI = getSP(props.context) ;
    return new Promise((resolve, reject) => {
      _sp.web.lists.getByTitle(formconst.REQUEST_LISTNAME).items.select("RefCode").filter(`Title eq '${officeName}'`)()
        .then((items) => {
          if (items.length > 0) {
            const officeRef = items[0].RefCode;
            console.log(officeRef)
            resolve(officeRef);
          } else {
            reject(new Error("Office not found"));
          }
        })
        .catch((error) => {
          reject(error);
        });
    });
  }


  export async function checklistFolderExistence(props:ITrpreqfrmProps,folderPath: string): Promise<boolean> {
   
    folderPath = folderPath.replace(formconst.BASE_URL, "");
      const _sp :SPFI = getSP(props.context) ;
      console.log(folderPath);
      const listfolder = await  _sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
      if(listfolder.Exists){
      return true;
      }else{return false;} // Folder exists
    
  }
  export async function checklibFolderExistence(props:ITrpreqfrmProps,folderPath: string): Promise<boolean> {
    folderPath = folderPath.replace(formconst.BASE_URL, "");
    console.log(folderPath)
    const _sp :SPFI = getSP(props.context) ;
    const listfolder = await _sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
    if(listfolder.Exists){
    return true;
    }else{return false;} // Folder exists
  
}

  

   
  

 
  
  
  
  
  
  
  
  
  