import * as React from 'react';
import styles from './DirectorySearchBox.module.scss';
import { IDirectorySearchBoxProps } from './IDirectorySearchBoxProps';

import { Fabric, TextField, SearchBox } from 'office-ui-fabric-react';
import { RxJsEventEmitter } from '../../../RxJsEventEmitter/RxJsEventEmitter';
import IEventData from '../../../RxJsEventEmitter/IEventData';

export default class DirectorySearchBox extends React.Component<IDirectorySearchBoxProps, {}> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();

  constructor(props: IDirectorySearchBoxProps) {
    super(props);
    this.state = {
      text: null,
    };
  }

  public render(): React.ReactElement<IDirectorySearchBoxProps> {
    return (
      <SearchBox
        styles={{ root: { width: 300 } }}
        placeholder={this.props.searchBoxPlaceholder}
        onChange={(_, newValue) => this._onChangeText(newValue)}
      />
    );
  }

  private _onChangeText = (
    text: string
  ): void => {
    var eventBody = {
      sharedData: text
    } as IEventData;

    this.eventEmitter.emit("filterTerms", eventBody);    
  }
}
