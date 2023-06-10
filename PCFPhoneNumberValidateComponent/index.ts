//import { IInputs, IOutputs } from "pcf-reactor-extension";
import {IInputs, IOutputs} from "./generated/ManifestTypes";
//import { useDataset, useControlContext } from "pcf-hooks";
import { TextField, ITextFieldProps } from '@fluentui/react/lib/TextField';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { isValidNumber, parsePhoneNumberFromString } from 'libphonenumber-js';

export class PCFPhoneNumberValidateComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private container: HTMLDivElement;
    private notifyOutputChanged: () => void;
    //private phoneNumber: string;
    //private countryCode: string;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
    
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view

        const props: ITextFieldProps = {
            placeholder: 'Enter phone number',
            value: context.parameters.phoneNumber.raw ? context.parameters.phoneNumber.raw : '',
            errorMessage: '',
            onChange: (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
                const phoneNumber = parsePhoneNumberFromString(newValue || '', context.parameters.countryCode.raw || '');

                if (phoneNumber && isValidNumber(phoneNumber)) {
                    context.parameters.phoneNumber.raw = phoneNumber.format("E.164");
                    props.errorMessage = '';
                } else {
                    context.parameters.phoneNumber.raw = '';
                    props.errorMessage = `Invalid phone number. Example format: ${context.parameters.countryCode.raw} XXX-XXX-XXXX`;
                }

                this.notifyOutputChanged(); // Notify the framework of output change

                ReactDOM.render(
                    React.createElement(TextField, props),
                    this.container
                );
            }
        };

        ReactDOM.render(
            React.createElement(TextField, props),
            this.container
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary

        ReactDOM.unmountComponentAtNode(this.container);
    }
}
