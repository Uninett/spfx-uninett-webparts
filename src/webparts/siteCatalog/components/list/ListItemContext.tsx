import * as React from 'react';
import { find } from 'core-js/library/fn/array';

// Object.assign polifyll
import { assign } from 'core-js/library/fn/object';
import { ContextualMenu, IContextualMenuItem, DirectionalHint } from 'office-ui-fabric-react/lib/ContextualMenu';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { deepEqual } from '../helpers/deepEqual';
import { T } from '../helpers/translate';
import * as strings from 'SiteCatalogWebPartStrings';
import { IContextFilter } from '../interfaces/iContextFilter';

//import { getUIResource, getSPResource } from './../../Resources/index';

export interface IState {
  selection?: { [key: string]: boolean };
  // isContextMenuVisible?: boolean;
  contextFilter?: Array<IContextFilter>;
  activeContextFilters?: Array<IContextFilter>;
  hideDialog: boolean;
}

export interface IProps {
  contextTarget: HTMLElement;
  isContextMenuVisible?: boolean;
  contextFilter?: Array<IContextFilter>;
  activeContextFilters?: Array<IContextFilter>;
  handleFiltered: (filteredInfo: Array<IContextFilter>, activeFilteredInfo: Array<IContextFilter>) => void;
  handleItemAction?: (actionType: ItemActionType) => void;
  handleDismiss: () => void;
  sortHandler: (sortedColumn: any) => void;
  setSortedColumn: (sortedColumn: any) => void;
  sortedColumn: any;
  selectedItem?: any;
}

export class ListItemContext extends React.Component<IProps, IState> {

  constructor(props: IProps) {
    super(props);

    this.state = {
      selection: {},
       hideDialog: true,
       activeContextFilters: []
      //, isContextMenuVisible: null // using NULL check on componentWillReceiveProps
    };
  }

  public shouldComponentUpdate(nextProps, nextState) {
    let { contextFilter: stateContextFilter, activeContextFilters: stateActiveContextFilter } = this.state;
    let { contextFilter: propContextFilter,  activeContextFilters: propActiveContextFilter} = this.props;

    if (
      deepEqual(nextState.contextFilter, stateContextFilter /* -> for nullCheck*/ || []) &&
      deepEqual(nextProps.contextFilter, propContextFilter /* -> for nullCheck*/ || []) &&
      deepEqual(nextState.activeContextFilters, stateActiveContextFilter /* -> for nullCheck*/ || []) &&
      deepEqual(nextProps.activeContextFilters, propActiveContextFilter /* -> for nullCheck*/ || []) &&
      nextState.hideDialog === this.state.hideDialog &&
      nextProps.isContextMenuVisible === this.props.isContextMenuVisible
    ) return false;

    return true;
  }

  public componentWillReceiveProps(nextProps) {
    let nextFieldName = nextProps.contextFilter && nextProps.contextFilter.length && nextProps.contextFilter[0].fieldName;
    let stateFieldName = this.state.contextFilter && this.state.contextFilter.length && this.state.contextFilter[0].fieldName;

    // contextFilter for different field
    if ((nextFieldName && nextFieldName !== stateFieldName) ) {
      this.setState({ contextFilter: nextProps.contextFilter, activeContextFilters: nextProps.activeContextFilters });
    }

    // if contextFilter is been cleared by MeasuresListItems
    if (!nextProps.contextFilter) this.setState({ contextFilter: null});
  }

  public render() {
    let { selection } = this.state;

    return (
      <div>

        {this.props.isContextMenuVisible ? (
          <ContextualMenu
            target={this.props.contextTarget}
            shouldFocusOnMount={false}
            isBeakVisible={true}
            directionalHint={DirectionalHint.bottomCenter}
            onDismiss={this._onDismiss}
            items={this.getContextItems()}
          />) : (null)}

      </div>
    );
  }

  private _showDialog = () => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = () => {
    this.setState({ hideDialog: true });
  }

  private handleFilterItemsClick = (ev, item?: IContextualMenuItem) => {
    //console.log('inside handleFilterItemsClick');
    /** nothing to handle if there are no contextFilter or MenuItem */
    if (!this.state.contextFilter || !item || !this.state.activeContextFilters) return;

    let newContextFilter = this.state.contextFilter.map(f => {
      // toggle isFiltered state based on clicked item
      return (f.fieldName === item.key && f.filterText === item.name) ? assign({}, f, { isFiltered: !f.isFiltered }) : f;
    });
    

    let newActiveFilter = [...this.state.activeContextFilters];
    let filterItem = find(this.state.contextFilter, f => {
      return f.fieldName === item.key && f.filterText === item.name;
    });

    if (this.state.activeContextFilters.length === 0) {
      if (filterItem !== null) newActiveFilter.push(filterItem);
    } else {
      
      !this.itemDoesExistInArray(newActiveFilter, item) 
        ? newActiveFilter.push(filterItem) 
        : newActiveFilter = newActiveFilter.filter(f => {
            if (f.fieldName !== item.key) return true;
            
            return (f.fieldName === item.key && f.filterText !== item.name);
          });

    }

    // let newActiveFilter = this.state.contextFilter.filter(f => {
    //   return (f.isFiltered == true);
    // });
    
    this.setState({ activeContextFilters: newActiveFilter, contextFilter: newContextFilter });
    this.notifyFilteredColumns(newActiveFilter, newContextFilter);
  }

  private getContextItems = () => {
    let { contextFilter, activeContextFilters } = this.state;

    if(!contextFilter) return;

    const isFiltered = (fieldName: string, filterText: string) => {

      for (var x = 0; x < activeContextFilters.length; x++) {
        if (activeContextFilters[x].fieldName == fieldName && activeContextFilters[x].filterText == filterText) {
          return true;
        }
      }

      return false;
    };

    // IContextualMenuItem devider object
    const devider: (key: string) => IContextualMenuItem = (key) => ({ key, name: '-' });

    // ClearFilter IContextualMenuItem
    const clearFilterItem: () => IContextualMenuItem = () => {
      if (!contextFilter) return;

      let clearFieldName = contextFilter && contextFilter.length && contextFilter[0].fieldName;
      //let isNoFilter = contextFilter ? (contextFilter.filter(filter => filter.isFiltered).length === 0) : true;

      let isNoFilter = activeContextFilters ? (activeContextFilters.filter(filter => filter.fieldName === clearFieldName).length === 0) : true;

      return {
        key: 'clearfilter-key',
        name: 'Clear filter',
        icon: 'ClearFilter',
        disabled: isNoFilter,
        onClick: () => {
          //debugger;
          if (!contextFilter) return;

          // set state with all isFilter sat to false
          const updatedContextFilter = contextFilter.map(filter => /* returns "filter" */ assign({}, filter, { isFiltered: false }));

          const updatedActiveFilter = activeContextFilters.filter(f => f.fieldName !== clearFieldName);
          this.setState({ contextFilter: updatedContextFilter, activeContextFilters: updatedActiveFilter });


          this.props.handleFiltered(updatedContextFilter, updatedActiveFilter);
        }

      };
    };

    if(contextFilter[0].fieldName === 'more') {
      let moreItems: IContextualMenuItem[] = [
        {
          key: 'SharepointSite', text: strings.Sharepoint, iconProps: { iconName: 'SharepointLogo' },
          canCheck: true, isChecked: false,
          onClick: (ev) => { 
            this.props.handleDismiss();
            this.props.handleItemAction('sharepoint');
          }
        },
        {
          key: 'PlannerSite', text: strings.Planner, iconProps: { iconName: 'PlannerLogo' },
          canCheck: true, isChecked: false,
          onClick: (ev) => { 
            this.props.handleDismiss();
            this.props.handleItemAction("planner");
          }
        }
        ,
        {
          key: 'EmailSite', text: strings.Email, iconProps: { iconName: 'OutlookLogo' },
          canCheck: true, isChecked: false,
          onClick: (ev) => { 
            this.props.handleDismiss();
            this.props.handleItemAction("email");
          }
        },
        {
          key: 'CalendarSite', text: strings.Calendar, iconProps: { iconName: 'Calendar' },
          canCheck: true, isChecked: false,
          onClick: (ev) => { 
            this.props.handleDismiss();
            this.props.handleItemAction("calendar");
          }
        }
      ];
      
      return moreItems;
    }

    let sortItems: IContextualMenuItem[] = [
      {
        key: 'SortAsc', name: strings.SortAsc, icon: 'SortUp',
        canCheck: true,
        isChecked: this.isSorted(this.state.contextFilter[0].fieldName, 'asc'),
        onClick: this._onToggleSelect
      },
      {
        key: 'SortDesc', name: strings.SortDesc, icon: 'SortDown',
        canCheck: true,
        isChecked: this.isSorted(this.state.contextFilter[0].fieldName, 'desc'),
        onClick: this._onToggleSelect
      },
    ];

    let filterItems: IContextualMenuItem[] = !contextFilter
      // empty array if there is no contextFilter
      ? []

      // else build ContextMenuItems from contextFilter
      : [clearFilterItem(), ...contextFilter.map((filter, i) => {
        return {
          key: filter.fieldName,
          name: filter.filterText,
          canCheck: true,
          isChecked: isFiltered(filter.fieldName, filter.filterText),
          onClick: this.handleFilterItemsClick
        };
      })];

    let contextItems: IContextualMenuItem[] = [
      ...sortItems,

      devider('dev1'),

      ...filterItems
    ];

    return contextItems;
  }

  private isSorted = (fieldName, direction):boolean => {

    

    if(!this.props.sortedColumn)
      return false;

    return this.props.sortedColumn.fieldName === fieldName && this.props.sortedColumn.direction === direction ? true : false;
  }


  private notifyFilteredColumns = (newActiveContextFilter: Array<IContextFilter>, newContextFilter: Array<IContextFilter>) => {
    this.props.handleFiltered(newContextFilter, newActiveContextFilter);
  }

  private _onToggleSelect = (ev?: React.MouseEvent<HTMLButtonElement>, item?: IContextualMenuItem) => {
    let { selection } = this.state;
    if (!selection || !item) return;

    selection[item.key] = !selection[item.key];


    let direction;

    if (item.key === 'SortAsc') {
      direction = 'asc';
    } else if (item.key === 'SortDesc') {
      direction = 'desc';
    }

    let sortedColumn = {
      fieldName: this.props.contextFilter[0].fieldName,
      direction: direction
    };

    // Remove sorting if user clicks on a sorting direction that is allready selected
    if (this.props.sortedColumn) {
      if (this.props.sortedColumn.fieldName === this.props.contextFilter[0].fieldName) {
        if (item.key === 'SortAsc' && this.props.sortedColumn.direction === 'asc') {
          sortedColumn = null;
        } else if (item.key === 'SortDesc' && this.props.sortedColumn.direction === 'desc') {
          sortedColumn = null;
        }
      }
    }

    this.props.sortHandler(sortedColumn);
    this.props.setSortedColumn(sortedColumn);
    this.props.handleDismiss();

    this.setState({
      selection: selection
    });
  }

  private _onClick = (event: React.MouseEvent<HTMLButtonElement>) => {
    this.props.handleDismiss();
  }

  private _onDismiss = (ev, dismissAll) => {
    if (dismissAll) return;
    this.props.handleDismiss();
  }

  private itemDoesExistInArray(newActiveFilter: IContextFilter[], item: IContextualMenuItem) {
    return find(newActiveFilter, f => (f.fieldName === item.key && f.filterText === item.name));
  }
}
