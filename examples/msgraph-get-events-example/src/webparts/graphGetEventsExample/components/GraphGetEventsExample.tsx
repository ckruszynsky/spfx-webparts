import * as React from 'react';
import styles from './GraphGetEventsExample.module.scss';
import { IGraphGetEventsExampleProps } from './IGraphGetEventsExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { List } from 'office-ui-fabric-react/lib/List';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {format} from 'date-fns';

export interface IGraphGetEventsExampleState {
  events:MicrosoftGraph.Event[]
}
export default class GraphGetEventsExample extends React.Component<IGraphGetEventsExampleProps, IGraphGetEventsExampleState> {
  constructor(props:IGraphGetEventsExampleProps){
    super(props);
    this.state = { events: []}
  }

  public async componentDidMount():Promise<void>{
    const events = await this.props.eventService.getEvents();
    this.setState({
      events: events
    });
  }


  public render(): React.ReactElement<IGraphGetEventsExampleProps> {
    return (
      <List items={this.state.events}
            onRenderCell={this._onRenderEventCell} />
    );
  }

  private _onRenderEventCell(item:MicrosoftGraph.Event) : JSX.Element {
    return (
        <div>
          <h3>{item.subject}</h3>
          <Icon iconName={'Calendar'} className={ styles.mailIcon} /> {'  '}
          {format( new Date(item.start.dateTime), 'MMMM DD, YYYY h:mm A')} - {format( new Date(item.end.dateTime), 'h:mm A')}
        </div>
    )
  }
}
