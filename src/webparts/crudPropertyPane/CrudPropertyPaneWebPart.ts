import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudPropertyPaneWebPart.module.scss';
import * as strings from 'CrudPropertyPaneWebPartStrings';

import {
  PropertyPaneContentSelector,
  IPropertyPaneContentSelectorProps
} from '../../controls/PropertyPaneContentSelector';

export interface ICrudPropertyPaneWebPartProps {
  description: string;
  workItem: string;
  myContinent: string;

}

export default class CrudPropertyPaneWebPart extends BaseClientSideWebPart<ICrudPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.crudPropertyPane} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Work Item: <strong>${escape(this.properties.workItem)}</strong></div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        <div>My Continent: <strong>${escape(this.properties.myContinent)}</strong></div>
        

        <div className={styles.buttons}>
          <button type="button" onClick={this.onOpenPanelClicked}>Add Task</button>
        </div>
      </div>
    </section>`;

      this.domElement.getElementsByTagName("button")[0]
    .addEventListener('click', (event: any) => {
      event.preventDefault();
      this.context.propertyPane.open();
    });
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('workItem', {
                  label: strings.WorkItemFieldLabel,
                  // onGetErrorMessage: this.validateWorkItem.bind(this)
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  // onGetErrorMessage: this.validateDescription.bind(this)
                }),
                new PropertyPaneContentSelector('myContinent', <IPropertyPaneContentSelectorProps> {
                  label: 'Continent where I currently Reside',
                  disabled:false,
                  selectedKey: this.properties.myContinent,
                  onPropertyChange: this.onContentSelectionChange.bind(this),
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  private onContentSelectionChange(propertyPath: string, newValue: any): void {
    const oldValue: any = this.properties[propertyPath];
    this.properties[propertyPath] = newValue;
    this.render();
  }


  //! Validation Controls

  // private validateDescription(textboxValue: string): string {
  //   const inputToValidate: string = textboxValue.toLowerCase();
  //   let len = {min:20};

  //   return (inputToValidate.length < len.min) ? 'Please Enter a more detailed Description' : "Valid Description length";
  //   // if (textboxValue == "") {
  //   //   return "Please Enter a more detailed Description";
  //   // } else {
  //   //   return "Valid Description length";
  //   // }
  // }

  // private validateWorkItem(textboxValue: string): string {
  //   const inputToValidate: string = textboxValue.toLowerCase();
    

  //   return (inputToValidate === null || inputToValidate === "") ? 'Please Enter Detailed Work Item Title' : "Work Item Title Entered";
  // }
}
