import * as React from 'react';
//import styles from './HelloWorldPriyanka.module.scss';
import { IGraphapiProps } from './IGraphapiProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from "@microsoft/sp-http";

//Email interface
interface IEmails{
subject: string;
weblink: string;
from:{
  emailAddress:{
    name: string;
    address: string;
  };
};
recivedDateTime: any;
bodypreview: string;
isread: any;
}

//All items interface
interface IAllItems{
  AllEmails:IEmails[];
}
export default class HelloWorldPriyanka extends React.Component<IGraphapiProps,IAllItems > {
  
  constructor(props:IGraphapiProps,state:IAllItems){
    super(props);
    this.state={
      AllEmails:[],
    };
  }
  componentDidMount(): void {
    this.getMyEmails;
  }
   public getMyEmails (){
    console.log("test emails");
    alert("hi");
   

    // below code is default code to get
    this.props.context.msGraphClientFactory //default syntax
      .getClient("3") //version updated to 3
      .then((client: MSGraphClientV3): void =>
      {
        client
          .api("me/messages") //to get messages
          .version("v1.0")
          .select("subject,webLink, from,receivedDateTime,isRead,bodyPreview") // selected columns from response preview
          .get((err: any, res: any) => {
            this.setState({
              AllEmails: res.value,
            });
            // console.log(this.state.AllEmails);         //checking
            /*  console.log(res);
            console.log(err); */
          });
      });
  
  }

  public render(): React.ReactElement<IGraphapiProps> {
    return(
      <div>
        {this.state.AllEmails.map((email)=>{
      
     return (
      <div>
        <p>{email.from.emailAddress.name}</p>
        <p>{email.subject}</p>
        <p>{email.recivedDateTime}</p>
        <p>{email.bodypreview}</p>
      </div>
     );
        })}
        </div>


    );
  }
}
