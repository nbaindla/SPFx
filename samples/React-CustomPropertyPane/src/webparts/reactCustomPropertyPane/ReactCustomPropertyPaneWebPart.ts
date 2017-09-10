import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IWebPartContext,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'reactCustomPropertyPaneStrings';
import ReactCustomPropertyPane from './components/ReactCustomPropertyPane';
import { IReactCustomPropertyPaneProps } from './components/IReactCustomPropertyPaneProps';
import { IReactCustomPropertyPaneWebPartProps } from './IReactCustomPropertyPaneWebPartProps';
import pnp from "sp-pnp-js";

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class ReactCustomPropertyPaneWebPart extends BaseClientSideWebPart<IReactCustomPropertyPaneWebPartProps> {
  private ddlListOptions: IPropertyPaneDropdownOption[] = [];
  private ddlListViewOptions: IPropertyPaneDropdownOption[] = [];
  private listsDropdownDisabled: boolean = true;
  private listViewsDropdownDisabled: boolean;

  private ddlLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private ddlLibraryViewOptions: IPropertyPaneDropdownOption[] = [];
  private libraryDropdownDisabled: boolean = true;
  private libraryViewsDropdownDisabled: boolean;
  
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });

      this.listViewsDropdownDisabled = true;
      this._getListTitles().then((response) => {
        response.forEach((list: any) => {
          this.ddlListOptions.push({ "key": list.Title, "text": list.Title });
        });
      });

      this.libraryViewsDropdownDisabled = true;
      this._getLibraryTitles().then((response) => {
        console.log(this.ddlLibraryOptions);
        response.forEach((library: any) => {
          this.ddlLibraryOptions.push({ "key": library.Title, "text": library.Title });
        });
      });

    });
  }

  public render(): void {
    const element: React.ReactElement<IReactCustomPropertyPaneProps> = React.createElement(
      ReactCustomPropertyPane,
      {
        title: this.properties.title,
        description: this.properties.description,
        listName: this.properties.listName,
        listViewName: this.properties.listViewName,
        libraryName: this.properties.libraryName,
        libraryViewName: this.properties.libraryViewName,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.ddlListOptions;
    this.listViewsDropdownDisabled = !this.properties.listName || !this.ddlListOptions;
    this.listsDropdownDisabled = false;
    this.context.propertyPane.refresh();
  }

  private _getListTitles(): Promise<any> {
    return pnp.sp.web.lists.filter('Hidden eq false and BaseType ne 1').select('Id', 'Title').get().then((r: any): Promise<any> => {
      return r;
    });
  }

  private _getLibraryTitles(): Promise<any> {
    return pnp.sp.web.lists.filter('Hidden eq false and BaseType eq 1').select('Id', 'Title').get().then((r: any): Promise<any> => {
      return r;
    });
  }

  private GetViewTitles(lisName: any): Promise<any> {
    return pnp.sp.web.lists.getByTitle(lisName).views.select('Id', 'Title').get().then((r: any): Promise<any> => {
      return r;
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Basic Details'
          },
          groups: [
            {
              groupName: 'Group Name',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                }),

              ]
            }
          ]
        },
        {
          header: {
            description: strings.PageDataSource
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.GroupList,
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.SelectList,
                  options: this.ddlListOptions,
                  disabled: false
                }),
                PropertyPaneDropdown('listViewName', {
                  label: strings.SelectView,
                  options: this.ddlListViewOptions,
                  disabled: this.listViewsDropdownDisabled
                })
              ]
            },
            {
              groupName: strings.GroupLibraries,
              groupFields: [
                PropertyPaneDropdown('libraryName', {
                  label: strings.SelectLibrary,
                  options: this.ddlLibraryOptions,
                  disabled: false
                }),
                PropertyPaneDropdown('libraryViewName', {
                  label: strings.SelectView,
                  options: this.ddlLibraryViewOptions,
                  disabled: this.libraryViewsDropdownDisabled
                })
              ]
            }
          ]
        },

      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

    if (propertyPath === 'listName' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      const previousItem: string = this.properties.listViewName;
      this.properties.listViewName = undefined;
      this.onPropertyPaneFieldChanged('listViewName', previousItem, this.properties.listViewName);
      this.listViewsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      let itemOptions: IPropertyPaneDropdownOption[] = [];

      this.GetViewTitles(newValue).then((response) => {
        response.forEach((list: any) => {
          itemOptions.push({ key: list.Title, text: list.Title });
        });
      }).then((): void => {
        this.ddlListViewOptions = itemOptions;
        this.context.propertyPane.refresh();
      });

    }
    else if (propertyPath === 'libraryName' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      const previousItem: string = this.properties.libraryViewName;
      this.properties.libraryViewName = undefined;
      this.onPropertyPaneFieldChanged('libraryName', previousItem, this.properties.libraryViewName);
      this.libraryViewsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      let itemOptions: IPropertyPaneDropdownOption[] = [];

      this.GetViewTitles(newValue).then((response) => {
        response.forEach((list: any) => {
          itemOptions.push({ key: list.Title, text: list.Title });
        });
      }).then((): void => {
         this.ddlLibraryViewOptions = itemOptions;
        this.context.propertyPane.refresh();
      });

    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }

  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
