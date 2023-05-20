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
      _sp.web.lists.getByTitle(formconst.CUSTOMER_LISTNAME).items.select("Reference").filter(`Title eq '${customerName}'`)()
        .then((items) => {
          if (items.length > 0) {
            const customerRef = items[0].Reference;
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
 
  
  
  
  
  
  
  
  
  