import * as React from 'react';
import styles from './SearchBox.module.scss';
import { ISearchBoxProps } from './ISearchBoxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Fabric, TextField, ShimmeredDetailsList, SelectionMode, DetailsListLayoutMode, ConstrainMode } from 'office-ui-fabric-react';
import { RxJsEventEmitter } from '../../../RxJsEventEmitter/RxJsEventEmitter';
import IEventData from '../../../RxJsEventEmitter/IEventData';

const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '400px'
  }
};

export default class SearchBox extends React.Component<ISearchBoxProps, {}> {

  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();

  constructor(props: ISearchBoxProps) {
    super(props);
    this.state = {
      text: null,
    };
  }

  public render(): React.ReactElement<ISearchBoxProps> {
    return (
      <Fabric>
        <div>
          <TextField label={"Filter by name or department:"} onChange={this._onChangeText} styles={controlStyles} />
        </div>
      </Fabric>
    );
  }


  private _onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    var eventBody = {
      sharedData: text
    } as IEventData;

    this.eventEmitter.emit("filterTerms", eventBody);
  }
}
