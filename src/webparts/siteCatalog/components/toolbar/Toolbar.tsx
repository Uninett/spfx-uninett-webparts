import * as React from 'react';
import * as strings from 'SiteCatalogWebPartStrings';
import styles from '../SiteCatalog.module.scss';
import { ActionButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

export interface IToolBarProps {
  navigateHandler: (ev, direction) => void;
  enablePrev: boolean;
  enableNext: boolean;
  handleNewClick: (e) => void;
  loading: boolean;
  currPage: string;
  totalPages: string;
  showNewButton: boolean;
}

export interface IToolBarState {
}

export class ToolBar extends React.Component<IToolBarProps, IToolBarState> {

  constructor(props) {
    super(props);

    this.state = {
    };
  }

  public componentWillMount(){
  }

  public render(): React.ReactElement<IToolBarProps> {

    let items:IContextualMenuItem[] = [];
    
    if (this.props.showNewButton) {
      items = [{
        name: strings.New,
        key: "New",
        iconProps: { 
          iconName: "Add" 
        },
        onClick: this.props.handleNewClick
      }];
    }

    let farItems:IContextualMenuItem[] = [];

    farItems.push({
      name: strings.Back,
      key: "Back",
      iconProps: { 
        iconName: "NavigateBack" 
      },
      onClick: this.backNavHandler,
      disabled: !this.props.enablePrev
    });

    farItems.push({
      name: this.props.currPage + '/' + this.props.totalPages,
      key: "Pagenumber"
    });

    farItems.push({
      name: strings.Next,
      key: "Next",
      iconProps: { 
        iconName: "NavigateBackMirrored" 
      },
      onClick: this.nextNavHandler,
      className: styles.floatIconRight,
      disabled: !this.props.enableNext
    });


    let style = !this.props.loading ? {borderBottom: "2px solid #eaeaea"} : {};
    
    return (
      <div style={style}>
        
        <CommandBar
          items={items}
          farItems={farItems}
        />
        {this.props.loading ? <ProgressIndicator className={'toolbarProgressIndicator'} /> : '' }
      </div>
    );
  }

  public backNavHandler = (ev) => {
    this.props.navigateHandler(ev, 'back');
  }

  public nextNavHandler = (ev) => {
    this.props.navigateHandler(ev, 'next');
  }

}
