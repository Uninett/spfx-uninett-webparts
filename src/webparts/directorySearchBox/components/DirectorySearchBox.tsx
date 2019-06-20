import * as React from 'react';
import styles from './DirectorySearchBox.module.scss';
import { IDirectorySearchBoxProps } from './IDirectorySearchBoxProps';

import { Fabric, TextField, SearchBox } from 'office-ui-fabric-react';
import { RxJsEventEmitter } from '../../../RxJsEventEmitter/RxJsEventEmitter';
import IEventData from '../../../RxJsEventEmitter/IEventData';

const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '400px'
  }
};

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
        onChange={newValue => this._onChangeText(newValue)}
      />
      /*
      <Fabric>
        <div>
          <TextField label={this.props.searchBoxLabel} onChange={this._onChangeText} styles={controlStyles} />
        </div>
      </Fabric>
      */
    );
  }


  private _onChangeText = (
    text: string
  ): void => {
    var eventBody = {
      sharedData: text
    } as IEventData;

    try {
      this.eventEmitter.emit("filterTerms", eventBody);
    } catch (error) {
      console.error("API not valid.");
    }
    
  }
}
