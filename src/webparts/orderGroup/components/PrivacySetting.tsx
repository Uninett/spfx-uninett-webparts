import * as React from 'react';
import * as $ from 'jquery';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IPrivacySettingState } from './IPrivacySettingState';
import * as strings from 'OrderGroupWebPartStrings';

export interface IPrivacySettingProps {
    onChange: (option: IDropdownOption, index?: number) => void;
}

class PrivacySetting extends React.Component<IPrivacySettingProps, IPrivacySettingState> {

    public render() {
        return (
            <div>
                <Dropdown
                    defaultSelectedKey='Closed'
                    label={strings.PrivacySetting}
                    options={
                        [
                            { key: 'Open', text: strings.PrivacyPublic },
                            { key: 'Closed', text: strings.PrivacyPrivate },
                        ]
                    }
                    onChange={() => this.props.onChange}
                    required={true}
                />
            </div>);
    }

}

export { PrivacySetting };