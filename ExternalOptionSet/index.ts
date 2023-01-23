import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { DropdownMenuItemType, IDropdownOption } from "@fluentui/react/lib/components/Dropdown/Dropdown.types";
import { ExternalOptionSetComponent, IExternalOptionSetComponent } from "./ExternalOptionSetComponent";
import { DAO } from "./DAO";
import { Option } from "./models/Option";

export class ExternalOptionSet implements ComponentFramework.ReactControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _DAO: DAO;
    private _options: Option[];
    private _dropdownOptions: IDropdownOption[];
    private _selectedOption: number | null;
    private _notifyOutputChanged: () => void;

    /**
     * Empty constructor.
     */
    constructor() {
        this._options = new Array<Option>();
        this._dropdownOptions = new Array<IDropdownOption>();
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary): void {
        this._context = context;
        this._DAO = new DAO((<any>this._context).page.getClientUrl());
        this._notifyOutputChanged = notifyOutputChanged;
        let table = this._context.parameters.table.raw!;
        let optionset = this._context.parameters.optionset.raw!;
        this.defineOptions(optionset, table);
    }

    private defineOptions(optionset: string, table: string) {

        let states = new Array<Option>();

        if (optionset == "statecode")
            this._options = this._DAO.RetrieveStateAttributeMetadata(table);
        else if (optionset == "statuscode") {
            states = this._DAO.RetrieveStateAttributeMetadata(table);
            this._options = this._DAO.RetrieveStatusAttributeMetadata(table);
        }
        else
            this._options = this._DAO.RetrievePicklistAttributeMetadata(table, optionset);

        if (optionset == "statuscode" && states.length > 0) {
            states.forEach(state_ => {
                this._dropdownOptions.push({ key: state_.Value, text: state_.Label, itemType: DropdownMenuItemType.Header });
                this._options.forEach(option_ => {
                    if (option_.State !== null && option_.State == state_.Value)
                        this._dropdownOptions.push({ key: option_.Value, text: option_.Label });
                });
                this._dropdownOptions.push({ key: '-999999999', text: '-', itemType: DropdownMenuItemType.Divider });
            });
        }
        else {
            this._options.forEach(option_ => {
                this._dropdownOptions.push({ key: option_.Value, text: option_.Label });
            });
        }
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this._selectedOption = this._context.parameters.whole_none.raw;
        return this.renderControl();
    }

    private renderControl(): React.ReactElement {
        // const dropdownControlledExampleOptions = [
        //     { key: 1, text: 'Fruits', itemType: DropdownMenuItemType.Header },
        //     { key: 4, text: 'Orange', disabled: true },
        //     { key: 5, text: 'Grape' },
        //     { key: 6, text: '-', itemType: DropdownMenuItemType.Divider },
        // ];
        const props: IExternalOptionSetComponent = {
            options: this._dropdownOptions,
            selectedOption: this._selectedOption,
            onChange: this.onChange
        };
        return React.createElement(
            ExternalOptionSetComponent, props
        );
    }

    private onChange = (newValue: number | null) => {
        this._selectedOption = newValue;
        this._notifyOutputChanged();
    };

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {
            whole_none: this._selectedOption !== null ? this._selectedOption : undefined
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
