import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: '100%', minWidth: '200px' } };

export interface IExternalOptionSetComponent {
  options: IDropdownOption<any>[],
  selectedOption: number | null,
  onChange: (value: number | null) => void
}

export const ExternalOptionSetComponent = React.memo((contract: IExternalOptionSetComponent): JSX.Element => {

  const _onSelectedChanged = (event: any, option?: IDropdownOption) => {
    let value = (option?.key == null || option?.key === -1) ? null : option?.key as number;
    contract.onChange(value);
  }

  return (
    <Dropdown
      selectedKey={contract.selectedOption ? contract.selectedOption : undefined}
      onChange={_onSelectedChanged}
      options={contract.options}
      styles={dropdownStyles}
    />);
}
);