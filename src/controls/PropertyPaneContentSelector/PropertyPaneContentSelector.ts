import * as React from "react";
import * as ReactDom from 'react-dom';
import {
    IPropertyPaneField,
    PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import {IDropdownOption} from 'office-ui-fabric-react/lib/components/Dropdown';
import { IContentSelectorProps } from "./components/IContentSelectorProps";
import ContentSelector from "./components/ContentSelector";
import {
    IPropertyPaneContentSelectorProps,
    IPropertyPaneContentSelectorInternalProps
} from './';

export class PropertyPaneContentSelector implements IPropertyPaneField<IPropertyPaneContentSelectorProps> {
    public type: PropertyPaneFieldType.Custom;
    public properties: IPropertyPaneContentSelectorInternalProps;
    private element: HTMLElement;

    constructor(public targetProperty: string, properties: IPropertyPaneContentSelectorProps) {
        this.properties = {
            key: properties.label,
            label: properties.label,
            disabled: properties.disabled,
            selectedKey: properties.selectedKey,
            onPropertyChange: properties.onPropertyChange,
            onRender: this.onRender.bind(this)
        };
    }

    public render(): void {
        if (!this.element) {
            return;
        }
    }

    private onRender(element: HTMLElement): void {
        if (!this.element) {
            this.element = element;
        }

        const reactElement: React.ReactElement<IContentSelectorProps> = React.createElement(ContentSelector, <IContentSelectorProps> {
            label: this.properties.label,
            onChanged: this.onChanged.bind(this),
            selectedKey: this.properties.selectedKey,
            disabled: this.properties.disabled,
            stateKey: new Date().toString()
        });
        ReactDom.render(reactElement, element);
    }

    private onChanged(option: IDropdownOption, index?: number): void {
        this.properties.onPropertyChange(this.targetProperty, option.key);
    }
}