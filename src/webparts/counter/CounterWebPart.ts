import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CounterWebPartStrings';

import styles from './CounterWebPart.module.scss';

import {sp} from '@pnp/sp/presets/all';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

export interface ICounterWebPartProps {
  title: string;
  selectedList: any;
  selectedField: any;
}

export interface IPropertyPaneDropdownOption{
  key: string;
  text: string;
}

export default class CounterWebPart extends BaseClientSideWebPart<ICounterWebPartProps> {
  private listDropDownOptions: IPropertyPaneDropdownOption[];
  private fieldsDropDownOptions: IPropertyPaneDropdownOption[];
  
  private _getListData(): Promise<any>{
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + 
      "/_api/web/lists/GetByTitle('" + this.properties.selectedList + "')/Items",
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) =>{
      return response.json();
    })
  }

  private _getListFields(selectedList: any):Promise<any>{
    return sp.web.lists.getByTitle(selectedList).fields.filter("ReadOnlyField eq false and Hidden eq false").get().then(function(result) {
      var fields = []
      result.forEach(element => {fields.push({key:element.EntityPropertyName, text: element.Title})});
      return fields;
    });
  }

  private _removeLists(element: any): any{
    const removableListTitles = ["Documents", "Form Templates", "Site Pages", "Style Library"];
    if(removableListTitles.indexOf(element.Title) < 0){
      return element;
    }
  }

  private _getLists():Promise<any>{
      return sp.web.lists.filter('Hidden eq false')
      .get()
      .then((data) =>{
        var filtered = data.filter(this._removeLists);
        console.log(filtered);
        return filtered;
      });
  }

  private _getGroupByCounter():Promise<any>{
    return sp.web.lists.getByTitle(this.properties.selectedList).items.select(this.properties.selectedField).get().then((items:any[]) =>{
      var x = [];
      items.forEach(element => {
        var flag = false
        for(var i=0; i<x.length; i++){
          if(element[this.properties.selectedField] === x[i].name){
            flag = true;
            x[i].value++;
          }
        }
        if(!flag)x.push({name:`${element[this.properties.selectedField]}`, value: 1});
      });
      return x;
    });
  }

  private _renderListCounter(): void{
    if(this.properties.selectedList){
      this._getListData()
      .then((response) => {
        let html: string = '<div>';
        //add amount of items
        html += `<p>${response.value.length}</p>`;
        html += '</div>';
        const mainCounterContainer: Element = this.domElement.querySelector('#mainCounterContainer');
        mainCounterContainer.innerHTML = html;
      });
    }
  }

  private _renderTitle(): void{
    const counterTitleContainer: Element = this.domElement.querySelector('#counterTitleContainer');

    let html: string = '<div>'
    html += `<p>${this.properties.title? this.properties.title : strings.Title}</p>`
    html += '</div>'

    counterTitleContainer.innerHTML = html;
  }

  private _renderGroupByCounter(): void{
    if(this.properties.selectedList && this.properties.selectedField){
      this._getGroupByCounter().then((response) =>{
        const groupByCounterContainer: Element = this.domElement.querySelector('#groupByCounterContainer');
        let html: string = '<ul>';
        response.forEach(element => {
          html += `<li><div><p>${element.name}</p><p>${element.value}</p></div></li>`;
        });
        html += '</ul>'
        groupByCounterContainer.innerHTML = html;
      });
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.counter}">
      <div class="${styles.counterTitleContainer}" id="counterTitleContainer">
      </div>
      <div class="${styles.mainCounterContainer}" id="mainCounterContainer">
      </div>
      <div class="${styles.groupByCounterContainer}" id="groupByCounterContainer">
      </div>
    </div>
    `
    this._renderTitle();
    this._renderListCounter();
    this._renderGroupByCounter();
  }

  protected async onInit(): Promise<void>{
    this.listDropDownOptions = [];
    this.fieldsDropDownOptions = [];
    const _ = await super.onInit();
    sp.setup({
      spfxContext: this.context
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart():void{
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this._getLists().then((response) =>{
      for(let i=0 ; i< response.length;i++){
        this.listDropDownOptions.push({key:response[i].Title,text:response[i].Title});
      }
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    });

    if(this.properties.selectedList){
      this._getListFields(this.properties.selectedList).then((response) =>{
        this.fieldsDropDownOptions = []
        response.forEach((element: any) => {
          this.fieldsDropDownOptions.push({key:element.key, text:element.text})
        });
        //this.properties.selectedField = this.fieldsDropDownOptions;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
    }
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
              groupFields:[
                PropertyPaneTextField('counterTitle', {
                  label:strings.Title
                })
              ]
            },
            {
              groupFields:[
                PropertyPaneDropdown('listDropdown',{
                  label:strings.SelectList,
                  options:this.listDropDownOptions
                })
              ],
            },
            {
              groupFields:[
                PropertyPaneDropdown('fieldsDropdown',{
                  label: 'Group by',
                  options:this.fieldsDropDownOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue:any, newValue:any):void{
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if(propertyPath ==='listDropdown' && newValue){
      this.properties.selectedList = newValue;
      this._getListFields(newValue).then((response) =>{
        const x = []
        response.forEach(element => {
          x.push({key:element.key, text:element.text})
        });
        this.fieldsDropDownOptions = x;
        this.context.propertyPane.refresh();
      });
    }
    else if(propertyPath === 'fieldsDropdown' && newValue){
      this.properties.selectedField = newValue;
    }
    else if(propertyPath==='counterTitle'){
      this.properties.title = newValue;
    }
  }
}
