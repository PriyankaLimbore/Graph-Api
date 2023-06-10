import * as React from 'react';
//import styles from './PriyaProfile.module.scss';
import { IPriyaProfileProps } from './IPriyaProfileProps';
//import { render } from 'react-dom';
//import { escape } from '@microsoft/sp-lodash-subset';
//import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClientV3 } from "@microsoft/sp-http";

interface Profile{
  givenname: string;
  jobTitle: string;
  mail: string;
  mobilephone: string;
  officelocation: string;
  preferredlanguage: string;
  surname: string;
  displayname: string;

}

export default class PriyaProfile extends React.Component<IPriyaProfileProps, Profile>{
  IPriyaProfileProps: any;
  constructor(props:IPriyaProfileProps,state:Profile){
    super(props);
    this.state = {
      givenname: "",
      jobTitle:"",
      mail:"",
      mobilephone:"",
      officelocation:"",
      preferredlanguage:"",
      surname:"",
      displayname:"",
};}

componentDidMount(): void {
  this.getmyprofile;
}

public getmyprofile(){
this.props.context.msGraphClientFactory //default syntax
    .getClient("3") //version updated to 3
    .then((client: MSGraphClientV3): void =>
    {
      client
        .api("/me") //to get messages
        .version("v1.0")
        .select("*") // selected columns from response preview
        .get((err: any, res: any) => {
          this.setState({
        displayname: res.displayname,
        mail:res.mail,
        jobTitle:res.jobtitle,
        givenname: res.givenname,
        mobilephone: res.mobilephone,
        officelocation: res.officelocation,
        preferredlanguage: res.prefferedlanguage,
        surname: res.surname,
        });
    });

}
public render(): React.ReactElement<IUserProfileProps> {
}

    return (
      <section>
      <div>
      <h1>
      <b>{this.state.displayname}</b>
      </h1>
      <p>Given Name : {this.state.givenname},</p>
      <p>Sur Name : {this.state.surname},</p>
      <p>Mail ID : {this.state.mail},</p>
      </div>
      </section>
    );
  }
}
