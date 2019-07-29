import * as React from 'react';
import * as strings from 'SiteCatalogWebPartStrings';
import styles from '../SiteCatalog.module.scss';
import { ActionButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';

export interface INavigationBarProps {
  navigateHandler: (ev, direction) => void;
  enablePrev: boolean;
  enableNext: boolean;
}

export interface INavigationBarState {
}

export class NavigationBar extends React.Component<INavigationBarProps, INavigationBarState> {

  constructor(props) {
    super(props);

    this.state = {
    };
  }

  public componentWillMount(){
  }

  public render(): React.ReactElement<INavigationBarProps> {
    
    return (
      <div>
        {this.props.enableNext || this.props.enablePrev ? 
          <div>
            {this.props.enablePrev ? 
              <ActionButton
                data-automation-id='test'
                iconProps={ { iconName: 'NavigateBack'} }
                disabled={ false }
                checked={ false }
                onClick={this.backNavHandler}
              >
                {strings.Back}
              </ActionButton>
              : ''}
            {this.props.enableNext ? 
              <ActionButton
                data-automation-id='test'
                iconProps={ { iconName: 'NavigateBackMirrored' } }
                disabled={ false }
                checked={ false }
                onClick={this.nextNavHandler}
              >
                {strings.Next}
              </ActionButton>
              : ''}
            
          </div>
          : ''}
        
      </div>
    );
  }

  public backNavHandler = (ev) => {
    console.log(ev);
    this.props.navigateHandler(ev, 'back');
  }

  public nextNavHandler = (ev) => {
    console.log(ev);
    this.props.navigateHandler(ev, 'next');
  }

}
