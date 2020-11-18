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
import {SPHttpClientResponse} from '@microsoft/sp-http';

export interface ICounterWebPartProps {
  title: string;
  selectedList: any;
  selectedView: any;
  selectedField: any;
}

export interface IPropertyPaneDropdownOption{
  key: string;
  text: string;
}

export default class CounterWebPart extends BaseClientSideWebPart<ICounterWebPartProps> {
  //All options for the drop down (lists, views, fields)
  private listDropDownOptions: IPropertyPaneDropdownOption[];
  private viewDropDownOptions: IPropertyPaneDropdownOption[];
  private fieldsDropDownOptions: IPropertyPaneDropdownOption[];
  
  //Get all the existing lists
  private _getLists():Promise<any>{
    return sp.web.lists.filter('Hidden eq false')
    .get()
    .then((data) =>{
      var filtered = data.filter(this._removeLists);
      return filtered;
    });
  }

  //Remove default lists (Documents, Form Templates, ...)
  private _removeLists(element: any): any{
    const removableListTitles = ["Documents", "Form Templates", "Site Pages", "Style Library"];
    if(removableListTitles.indexOf(element.Title) < 0){
      return element;
    }
  }

  // Get the data from a sepcific view
  private _getViewData(viewName: string):Promise<any>{
    return this._getViewQueryForList(this.properties.selectedList,viewName).then((res:any) => {
      return this._getItemsByViewQuery(this.properties.selectedList,res).then((items:any[])=>{
        const x = [];
          items.forEach((item:any) => {
              x.push(item);
          });
          return x;
      });
    }).catch(console.error);
  }

  // Get data from view (if no view is selected => default view is used)
  private _getListViewData():Promise<any>{
    if(this.properties.selectedView){
      return this._getViewData(this.properties.selectedView).then((response) => {
        return response;
      });
    }
    else{
      return sp.web.lists.getByTitle(this.properties.selectedList).defaultView().then((response) => {
        return this._getViewData(response.Title).then((resp) => {
          return resp;
        });
      });
    }
  }
  //First method that retrieves the View Query
  private _getViewQueryForList(listName:string,viewName:string):Promise<any> {
    let listViewData = "";
    if(listName && viewName){
        return sp.web.lists.getByTitle(listName).views.getByTitle(viewName).select("ViewQuery").get().then(v => {
            return v.ViewQuery;
        });
    } else {
        console.log('Data insufficient!');
        listViewData = "Error";
    }
  }
  //Second method that retrieves the View data based on the View Query and List name
  private _getItemsByViewQuery(listName:string, query:string):Promise<any> {
    const xml = '<View><Query>' + query + '</Query></View>';  
    return sp.web.lists.getByTitle(listName).getItemsByCAMLQuery({'ViewXml':xml}).then((res:SPHttpClientResponse) => {
        return res;
    });
  }

  //Get all the views of a list
  private _getListViews(): Promise<any>{
    return sp.web.lists.getByTitle(this.properties.selectedList).views().then((result) => {
      var views = [];
      result.forEach(element => {
        views.push({key: element.Title, text: element.Title});
      });
      return views;
    });
  }

  //Get all the visible fields in a list
  private _getListFields(selectedList: any):Promise<any>{
    return sp.web.lists.getByTitle(selectedList).fields.filter("ReadOnlyField eq false and Hidden eq false").get().then((result) => {
      var fields = [];
      result.forEach(element => {fields.push({key:element.EntityPropertyName, text: element.Title});});
      return fields;
    });
  }

  //Get amount of items for each different value in a selected field
  private _getGroupByCounter():Promise<any>{
    return this._getListViewData().then((items:any[]) =>{
      var x = [];
      let isBool = false;
      if(typeof items[0][this.properties.selectedField] === 'boolean'){
        x.push({name:`${this.properties.selectedField}: yes`, value: 0});
        x.push({name:`${this.properties.selectedField}: no`, value: 0});
        isBool = true;
      } 
      items.forEach(element => {
        if(isBool){
          if(element[this.properties.selectedField]){
            x[0].value++;
          }
          else{
            x[1].value++;
          }
        }
        else{
          var flag = false;
          for(var i=0; i<x.length; i++){
            if("" + element[this.properties.selectedField] === x[i].name){
              flag = true;
              x[i].value++;
            }
          }
          if(!flag)x.push({name:`${element[this.properties.selectedField]}`, value: 1});
        }
      });
      return x;
    });
  }

  //Render the main counter
  private _renderListCounter(): void{
    if(this.properties.selectedList){
      this._getListViewData()
      .then((response) => {
        let html: string = '<div>';
        //add amount of items
        html += `<p>${response.length}</p>`;
        html += '</div>';
        const mainCounterContainer: Element = this.domElement.querySelector('#mainCounterContainer');
        mainCounterContainer.innerHTML = html;
      });
    }
  }

  //Render the title of the web part
  private _renderTitle(): void{
    const counterTitleContainer: Element = this.domElement.querySelector('#counterTitleContainer');

    let html: string = '<div>';
    html += `<p>${this.properties.title? this.properties.title : strings.Title}</p>`;
    html += '</div>';

    counterTitleContainer.innerHTML = html;
  }

  //Render the counter for each unique value in a selected field
  private _renderGroupByCounter(): void{
    if(this.properties.selectedList && this.properties.selectedField){
      this._getGroupByCounter().then((response) =>{
        const groupByCounterContainer: Element = this.domElement.querySelector('#groupByCounterContainer');
        let html: string = '<ul>';
        response.forEach(element => {
          html += `<li><div><p>${element.name}</p><p>${element.value}</p></div></li>`;
        });
        html += '</ul>';
        groupByCounterContainer.innerHTML = html;
      });
    }
  }

  //General redner methode
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
    `;
    this._renderTitle();
    this._renderListCounter();
    this._renderGroupByCounter();
  }

  protected async onInit(): Promise<void>{
    this.listDropDownOptions = [];
    this.fieldsDropDownOptions = [];
    this.viewDropDownOptions = [];
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
    this._getLists().then((response) =>{
      for(let i=0 ; i< response.length;i++){
        this.listDropDownOptions.push({key:response[i].Title,text:response[i].Title});
      }
      this.context.propertyPane.refresh();
      this.render();
    });

    if(this.properties.selectedList){
      this._getListViews().then((response) =>{
        this.viewDropDownOptions = [];
        response.forEach((element: any) => {
          this.viewDropDownOptions.push({key:element.key, text: element.text});
        });
        this.context.propertyPane.refresh();
        this.render();
      });
      
      this._getListFields(this.properties.selectedList).then((response) =>{
        this.fieldsDropDownOptions = [];
        response.forEach((element: any) => {
          this.fieldsDropDownOptions.push({key:element.key, text:element.text});
        });
        this.context.propertyPane.refresh();
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
                PropertyPaneDropdown('viewsDropdown', {
                  label:strings.SelectView,
                  options:this.viewDropDownOptions
                })
              ]
            },
            {
              groupFields:[
                PropertyPaneDropdown('fieldsDropdown',{
                  label:strings.GroupBy,
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
    switch(propertyPath){
      case 'listDropdown':
        this.properties.selectedList = newValue;
        this._getListFields(newValue).then((response) =>{
          this.fieldsDropDownOptions = [];
          response.forEach((element: { key: any; text: any; }) => {
            this.fieldsDropDownOptions.push({key:element.key, text:element.text});
          });
          this.context.propertyPane.refresh();
        });
        this._getListViews().then((response: any) => {
          this.viewDropDownOptions = [];
          response.forEach((element: { key: any; text: any; }) => {
            this.viewDropDownOptions.push({key:element.key, text:element.text});
          });
          this.context.propertyPane.refresh();
        });
        break;
      case 'fieldsDropdown':
        this.properties.selectedField = newValue;
        break;
      case 'counterTitle':
        this.properties.title = newValue;
        break;
      case 'viewsDropdown':
        this.properties.selectedView = newValue;
        break;
    }
  }
}
