import * as React from 'react';
import {
  SearchBox
} from 'office-ui-fabric-react/lib/SearchBox';
import styles from '../SiteCatalog.module.scss';
import * as strings from 'SiteCatalogWebPartStrings';

export interface ISearchBoxContainerState {
  keyWord: string;
}

export interface ISearchBoxContainerProps {
  searchHandler: any;
  clearHandler: any;
}

// tslint:disable:jsx-no-lambda
export class SearchBoxContainer extends React.Component<ISearchBoxContainerProps, ISearchBoxContainerState> {

  public constructor(props: ISearchBoxContainerProps) {
    super(props);

    this.state = {
      keyWord: ''
    };
  }

  public render(): JSX.Element {
    return (
      <div className={styles["msSearchBoxContainer"]}>
        <SearchBox
          placeholder={strings.SearchPlaceHolder}
          onClear={ (ev) => {
            this.props.clearHandler();
          } }
          onChange={ (_, newValue) => { this.setState({ keyWord: newValue }); } }
          onSearch={ (newValue) => { this.props.searchHandler(newValue); } }
        />
      </div>
    );
  }

}