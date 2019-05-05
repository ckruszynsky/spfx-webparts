import * as React from 'react';
import styles from './GraphMailMessageExample.module.scss';
import { IGraphMailMessageExampleProps } from './IGraphMailMessageExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MailMessage } from '../../../models';
import {IGraphMailMessageExampleState} from './IGraphMailMessageExampleState';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
export default class GraphMailMessageExample extends React.Component<IGraphMailMessageExampleProps, IGraphMailMessageExampleState> {
  constructor(props:IGraphMailMessageExampleProps){
    super(props);
    this.state = {
      messages:new Array<MailMessage>()
    }
  }

  public async componentDidMount():Promise<void>{
    const messages = await this.props.mailService.getMail();
    this.setState({messages:messages});
  }

  public render(): React.ReactElement<IGraphMailMessageExampleProps> {
    let content:JSX.Element = <div>{JSON.stringify(this.state.messages)}</div>;    

    return (
      <div className={styles.graphMailMessageExample}>
      <FocusZone direction={FocusZoneDirection.vertical}>
        <List items={this.state.messages} 
          onRenderCell={this._onRenderCell} 
          />
      </FocusZone>
      </div>
    );
  }

  private _onRenderCell(item: any, index: number | undefined): JSX.Element {
    return (
      <div className={styles.itemCell} data-is-focusable={ true }>     
        
        <Icon iconName={'Mail'} className={ styles.mailIcon} />                
        <div className={styles.itemContent}>
          <div className={styles.itemName}>{ item.subject }</div>          
          <div>From: { item.sender.emailAddress.name }</div>
          <p>{item.receivedDateTime}</p>
          <div>{item.bodyPreview}</div>
        </div>

      </div>
    );
  }
}

