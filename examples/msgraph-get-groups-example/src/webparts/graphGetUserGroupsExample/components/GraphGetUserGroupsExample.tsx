import * as React from 'react';
import styles from './GraphGetUserGroupsExample.module.scss';
import { IGraphGetUserGroupsExampleProps } from './IGraphGetUserGroupsExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Customizer } from 'office-ui-fabric-react';
import { FluentCustomizations } from '@uifabric/fluent-theme';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { getTheme } from 'office-ui-fabric-react/lib/Styling';

import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardDetails,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  DocumentCardType,
  IDocumentCardActivityPerson
} from 'office-ui-fabric-react/lib/DocumentCard';
import { Stack, StackItem } from 'office-ui-fabric-react/lib/Stack';

export interface IGraphGetUserGroupsExampleState {
  groups: MicrosoftGraph.Group[]
}

export default class GraphGetUserGroupsExample extends React.Component<IGraphGetUserGroupsExampleProps, IGraphGetUserGroupsExampleState> {


  constructor(props: IGraphGetUserGroupsExampleProps) {
    super(props);
    this.state = {
      groups: []
    };
  }

  public async componentDidMount(): Promise<void> {
    const groups = await this.props.service.getGroups();
    this.setState({ groups: groups });
  }

  public render(): React.ReactElement<IGraphGetUserGroupsExampleProps> {
    const theme = getTheme();
    const previewOutlookUsingIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewIconProps: {
            iconName: 'SharePointLogo',
            styles: {
              root: {
                fontSize: 42,
                color: '#0078d7',
                backgroundColor: theme.palette.neutralLighterAlt
              }
            }
          },
          width: 144
        }
      ],
      styles: {
        previewIcon: { backgroundColor: theme.palette.neutralLighterAlt }
      }
    };
    return (
      <Customizer {...FluentCustomizations}>
        <Stack gap={10}>
          {this.state.groups.map(group => 
          <StackItem grow={1} >
          <DocumentCard type={DocumentCardType.normal}>
            <DocumentCardPreview {...previewOutlookUsingIcon} />
            <DocumentCardDetails>
              <DocumentCardTitle title={group.displayName} shouldTruncate={true} />
              <p style={{paddingLeft:'25px'}}>{group.description}</p>
            </DocumentCardDetails>
          </DocumentCard>
          </StackItem>
          )}
        </Stack>
      </Customizer>
    );
  }
}
