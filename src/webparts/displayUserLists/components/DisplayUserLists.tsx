import * as React from 'react';
import styles from './DisplayUserLists.module.scss';

import { escape } from '@microsoft/sp-lodash-subset';

import {IDisplayUserListsState} from './IDisplayUserListsState';
import {IDisplayUserListsProps} from './IDisplayUserListsProps';
import {IUserItem} from './IUserItem';

import {
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/components/Persona';
import { Link } from 'office-ui-fabric-react/lib/components/Link';

export default class DisplayUserLists extends React.Component<IDisplayUserListsProps, IDisplayUserListsState> {

  constructor(props: IDisplayUserListsProps, state: IDisplayUserListsState) {
    super(props);

    // Initialize the state of the component
    this.state = {
      users: []      
    };
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api('users')      
      .get((error: any, res: any, rawResponse?: any) => {
        if (error) {
          console.error(error);
          return;
        }

        // Prepare the output array
        var users: Array<IUserItem> = new Array<IUserItem>();

        // Map the JSON response to the output array
        res.value.map((item: any) => {
          users.push( { 
            displayName: item.displayName,
            mail: item.mail,
            userPrincipalName: item.userPrincipalName,
            phone : item.businessPhones[0]
          });
        });

        // Update the component state accordingly to the result
        this.setState(
          {
            users: users,
          }
        );
      });
  
    
  }


  public render(): React.ReactElement<IDisplayUserListsProps> {
    const items = this.state.users.map((item,key) =>{
       return (<Persona primaryText={item.displayName}
              secondaryText={item.mail}              
              tertiaryText={item.phone}              
              imageUrl={""}
              size={PersonaSize.size48} />);
      
    });
    return (
        <div>{items}</div>
    );
  }


}
