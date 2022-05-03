import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  PropertyPaneHorizontalRule,
  PropertyPaneDropdown,
  PropertyPaneButton,
  PropertyPaneButtonType,


} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { INavLink } from 'office-ui-fabric-react/lib/Nav';

import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

//Copied from AdvancedPagePropertiesWebPart.ts
import * as _lodashAPP from 'lodash';


import { SPService } from '../../Service/SPService';

import * as strings from 'FpsPageInfoWebPartStrings';
import FpsPageInfo from './components/FpsPageInfo';
import { IFpsPageInfoProps } from './components/IFpsPageInfoProps';


import { Log } from './components/AdvPageProps/utilities/Log';

export interface IFpsPageInfoWebPartProps {
  description: string;

  //Copied from AdvancedPagePropertiesWebPart.ts
  title: string;
  selectedProperties: string[];
}

export default class FpsPageInfoWebPart extends BaseClientSideWebPart<IFpsPageInfoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  //Added from react-page-navigator component
  private anchorLinks: INavLink[] = [];

  //Copied from AdvancedPagePropertiesWebPart.ts
  private availableProperties: IPropertyPaneDropdownOption[] = [];


  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then(async _ => {

      // other init code may be present

      let mess = 'onInit - ONINIT: ' + new Date().toLocaleTimeString();
      console.log(mess);

      //https://stackoverflow.com/questions/52010321/sharepoint-online-full-width-page
      if ( window.location.href &&  
        window.location.href.toLowerCase().indexOf("layouts/15/workbench.aspx") > 0 ) {
          
        if (document.getElementById("workbenchPageContent")) {
          document.getElementById("workbenchPageContent").style.maxWidth = "none";
        }
      } 

      //console.log('window.location',window.location);
      sp.setup({
        spfxContext: this.context
      });

      //Added from react-page-navigator component
      this.anchorLinks = await SPService.GetAnchorLinks(this.context);

      //Have to insure selectedProperties always is an array from AdvancedPagePropertiesWebPart.ts
      // if ( !this.properties.selectedProperties ) { this.properties.selectedProperties = []; }

    });
  }

  public render(): void {
    const element: React.ReactElement<IFpsPageInfoProps> = React.createElement(
      FpsPageInfo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,

        pageNavigator:   {
          description: 'desc passed from main web part',
          anchorLinks: this.anchorLinks,
        }
        
      }
    );

    ReactDom.render(element, this.domElement);
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

  //Copied from AdvancedPagePropertiesWebPart.ts
  private async getPageProperties(): Promise<void> {
    Log.Write("Getting Site Page fields...");
    const list = await sp.web.lists.ensureSitePagesLibrary();
    const fi = await list.fields();

    this.availableProperties = [];
    Log.Write(`${fi.length.toString()} fields retrieved!`);
    fi.forEach((f) => {
      if (!f.FromBaseType && !f.Hidden && f.SchemaXml.indexOf("ShowInListSettings=\"FALSE\"") === -1
          && f.TypeAsString !== "Boolean" && f.TypeAsString !== "Note") {
        const internalFieldName = f.InternalName == "LinkTitle" ? "Title" : f.InternalName;
        this.availableProperties.push({ key: internalFieldName, text: f.Title });
        Log.Write(f.TypeAsString);
      }
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected onAddButtonClick (value: any) {
    this.properties.selectedProperties.push(this.availableProperties[0].key.toString());
  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected onDeleteButtonClick (value: any) {
    Log.Write(value.toString());
    var removed = this.properties.selectedProperties.splice(value, 1);
    Log.Write(`${removed[0]} removed.`);
  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath.indexOf("selectedProperty") >= 0) {
      Log.Write('Selected Property identified');
      let index: number = _lodashAPP.toInteger(propertyPath.replace("selectedProperty", ""));
      this.properties.selectedProperties[index] = newValue;
    }
  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    Log.Write(`onPropertyPaneConfigurationStart`);
    await this.getPageProperties();
    this.context.propertyPane.refresh();
  }






  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    Log.Write(`getPropertyPaneConfiguration`);

    // Initialize with the Title entry
    var propDrops: IPropertyPaneField<any>[] = [];
    propDrops.push(PropertyPaneTextField('title', {
      label: strings.TitleFieldLabel
    }));
    propDrops.push(PropertyPaneHorizontalRule());
    // Determine how many page property dropdowns we currently have
    this.properties.selectedProperties.forEach((prop, index) => {
      propDrops.push(PropertyPaneDropdown(`selectedProperty${index.toString()}`,
        {
          label: strings.SelectedPropertiesFieldLabel,
          options: this.availableProperties,
          selectedKey: prop,
        }));
      // Every drop down gets its own delete button
      propDrops.push(PropertyPaneButton(`deleteButton${index.toString()}`,
      {
        text: strings.PropPaneDeleteButtonText,
        buttonType: PropertyPaneButtonType.Command,
        icon: "RecycleBin",
        onClick: this.onDeleteButtonClick.bind(this, index)
      }));
      propDrops.push(PropertyPaneHorizontalRule());
    });
    // Always have the Add button
    propDrops.push(PropertyPaneButton('addButton',
    {
      text: strings.PropPaneAddButtonText,
      buttonType: PropertyPaneButtonType.Command,
      icon: "CirclePlus",
      onClick: this.onAddButtonClick.bind(this)
    }));

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }, //End this group
            {
              groupName: strings.SelectionGroupName,
              groupFields: propDrops
            }
          ]
        }
      ]
    };
  }
}
