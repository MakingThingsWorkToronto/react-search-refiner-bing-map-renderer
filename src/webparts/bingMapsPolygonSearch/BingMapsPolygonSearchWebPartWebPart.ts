
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Text, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'BingMapsSearchWebPartStrings';
import BingMaps from '../../components/BingMap/BingMap';
import IBingMapProps from '../../components/BingMap/IBingMapProps';
import IBingMapsPolygonSearchWebPartWebPartProps from './IBingMapsPolygonSearchWebPartWebPartProps';
import { ResultService, ISearchEvent} from '../../services/ResultService/ResultService';
import { ISearchResults, ISearchResult } from '../../models/ISearchResult';
import IResultService from '../../services/ResultService/IResultService';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';
import TemplateService from '../../services/TemplateService/TemplateService';
import MockTemplateService from '../../services/TemplateService/MockTemplateService';
import BingMap from '../../components/BingMap/BingMap';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';


export default class BingMapsPolygonSearchWebPartWebPart extends BaseClientSideWebPart<IBingMapsPolygonSearchWebPartWebPartProps> {

  private _searchResults : ISearchResults;
  private _resultService: IResultService;
  private _map: React.ReactElement<IBingMapProps>;
  private _component : BingMaps;
  private FIELDS: string = "Title,Description,Location";
  private _propertyFieldCodeEditor = null;
  private _propertyFieldCodeEditorLanguages = null;
  private _templateService: BaseTemplateService;
  private _mapTypeIds : IPropertyPaneDropdownOption[] = [
    {key:"aerial",text:"Aerial"},
    {key:"canvasDark",text:"Dark"},
    {key:"canvasLight",text:"Light"},
    {key:"birdseye",text:"Birdseye"},
    {key:"grayscale",text:"Grayscale"},
    {key:"ordnanceSurvey",text:"Ordnance Survey (UK Only)"},
    {key:"road",text:"Road"},
    {key:"streetside",text:"Street Side"}
  ];
  
  protected onInit(): Promise<void> {
    
    if (Environment.type === EnvironmentType.Local) {
      this._templateService = new MockTemplateService(this.context.pageContext.cultureInfo.currentUICultureName);
    } else {
        this._templateService = new TemplateService(this.context.spHttpClient, this.context.pageContext.cultureInfo.currentUICultureName);
    }

    this._resultService = new ResultService();
    this._resultService.registerRenderer(
            this.componentId, 
            'Bing Maps', 
            'MapPin', 
            (e) => this.onChangeHappened(e), this.properties.columns.map((i) => { return i.name; }) as string[]
    );

    return Promise.resolve();
    
  }

  private setMockResults() : void {}

  public render(): void {

    if (Environment.type === EnvironmentType.Local) {
      this.setMockResults();
    }
    
    let center: number[] = [ this.tryParseFloat(this.properties.centerLatitude), this.tryParseFloat(this.properties.centerLongitude) ];
    this._map = React.createElement(
      BingMaps, { 
        componentId: this.componentId, 
        polygonResults: this._searchResults,
        templateService: this._templateService,
        bingMapsAPIKey: this.properties.bingMapsAPIKey,
        mapTypeId: this.properties.mapTypeId,
        zoom: this.properties.zoom,
        navigationBarMode: this.properties.navigationBarMode,
        supportedMapTypes: this.properties.supportedMapTypes,
        columns: this.properties.columns,
        titleColumnName: this.properties.showLabels == false ? "" : this.properties.titleColumnName,
        polygonColumnName: this.properties.polygonColumnName,
        targetColumnName: this.properties.targetColumnName,
        fillColor: this.toColorString(this.properties.fillColor),
        strokeColor: this.toColorString(this.properties.strokeColor),
        strokeThickness: this.properties.strokeThickness,
        center: center
      }
    );

    this._component = ReactDom.render(this._map, this.domElement) as BingMaps;
    
  }

  private toColorString(val:any) : string {
    if(!val) return val;
    if(typeof val == "string") return val;
    if(typeof val == "object") {
      return 'rgba(' + val.r.toString() + ',' + val.g.toString() + ',' + val.b.toString() + ',' + val.a.toString() + ')';
    }
  }

  private tryParseFloat(val: string) : number {
    try {
      return parseFloat(val);
    } catch(ex){
      return 0;
    }
  }

  public onChangeHappened(e: ISearchEvent) {
    if(this._map) this._map.props.polygonResults = e.results;
    if(this._component) this._component.forceUpdate();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneConfigurationStart() {
    await this.loadPropertyPaneResources();
  }

  protected async loadPropertyPaneResources(): Promise<void> {

    // tslint:disable-next-line:no-shadowed-variable
    const { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } = await import(
        /* webpackChunkName: 'search-property-pane' */
        '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );

    this._propertyFieldCodeEditor = PropertyFieldCodeEditor;
    this._propertyFieldCodeEditorLanguages = PropertyFieldCodeEditorLanguages;
      
  }

  /**
     * Checks if a field if empty or not
     * @param value the value to check
     */
    private _validateEmptyField(value: string): string {

      if (!value) {
          return "This is a required field.";
      }

      return '';
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
              groupName: strings.bingMapsGroupNameLabel,
              groupFields: [
                PropertyPaneTextField('bingMapsAPIKey', {
                  label: strings.bingMapsAPIKeyLabel,
                  value: this.properties.bingMapsAPIKey,
                  onGetErrorMessage: this._validateEmptyField.bind(this)
                }),
                PropertyFieldNumber('zoom', {
                  key: "zoom",
                  label: strings.zoomLabel,
                  value: this.properties.zoom,
                  onGetErrorMessage: this._validateEmptyField.bind(this),
                  minValue: 1,
                  maxValue: 19
                }),
                PropertyPaneDropdown('mapTypeId', {
                  label: strings.mapTypeIdLabel,
                  selectedKey: this.properties.mapTypeId,
                  options: this._mapTypeIds
                }),
                PropertyFieldMultiSelect('supportedMapTypes', {
                  key: 'supportedMapTypes',
                  label: strings.supportedMapTypesLabel,
                  options: this._mapTypeIds,
                  selectedKeys: this.properties.supportedMapTypes
                }),
                PropertyPaneTextField('centerLatitude', {
                  label: strings.centerLatitudeLabel,
                  value: this.properties.centerLatitude.toString(),
                  onGetErrorMessage: this._validateEmptyField.bind(this),
                }),
                PropertyPaneTextField('centerLongitude', {
                  label: strings.centerLongitudeLabel,
                  value: this.properties.centerLongitude.toString(),
                  onGetErrorMessage: this._validateEmptyField.bind(this)
                }),
              ]
            },
            {
              groupName: strings.columnConfigurationGroupLabel,
              groupFields: [
                PropertyPaneTextField('titleColumnName', {
                  label: strings.titleColumnNameLabel,
                  value: this.properties.titleColumnName,
                  onGetErrorMessage: this._validateEmptyField.bind(this)
                }),
                PropertyPaneTextField('polygonColumnName', {
                  label: strings.polygonColumnNameLabel,
                  value: this.properties.polygonColumnName
                }),
                PropertyPaneTextField('targetColumnName', {
                  label: strings.targetColumnNameLabel,
                  value: this.properties.targetColumnName
                }),
                PropertyFieldCollectionData('columns',{
                  key: 'columns',
                  label: strings.columnsLabel,
                  panelHeader: strings.columnsPanelHeader,
                  manageBtnLabel: strings.columnsButtonLabel,
                  value: this.properties.columns,
                  enableSorting: true,
                  disableItemCreation: true,
                  disableItemDeletion: true,
                  fields: [{
                      id:"name",
                      title:strings.columnsNameColumnTitle,
                      type: CustomCollectionFieldType.string,
                      required: true
                    }]
                })
              ]
            },
            {
              groupName: strings.colorsStylesGroupLabel,
              groupFields: [
                PropertyPaneToggle('showLabels', {
                    label: strings.showLabelsLabel, 
                    checked: this.properties.showLabels
                }), 
                PropertyFieldNumber('strokeThickness', {
                  key: "strokeThickness",
                  label: strings.strokeThicknessLabel,
                  value: this.properties.strokeThickness,
                  onGetErrorMessage: this._validateEmptyField.bind(this),
                  minValue: 1,
                  maxValue: 10
                }),
                PropertyFieldColorPicker('fillColor', {
                  label: strings.fillColorLabel,
                  properties: this.properties,
                  selectedColor: this.properties.fillColor,
                  onPropertyChange: (e) => {},
                  style: PropertyFieldColorPickerStyle.Full,
                  valueAsObject: true,
                  iconName: 'BucketColor',
                  key: 'fillColorID'
                }),
                PropertyFieldColorPicker('strokeColor', {
                  label: strings.strokeColorLabel,
                  properties: this.properties,
                  selectedColor: this.properties.strokeColor,
                  onPropertyChange: (e)=> {},
                  style: PropertyFieldColorPickerStyle.Full,
                  valueAsObject: true,
                  iconName: 'Line',
                  key: 'strokeColorID'
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
