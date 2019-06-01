import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Text, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'BingMapsSearchWebPartStrings';
import BingMaps from '../../components/BingMap/BingMap';
import IBingMapProps from '../../components/BingMap/IBingMapProps';
import IBingMapsSearchWebPartProps from './IBingMapsSearchWebPartProps';
import { ResultService, ISearchEvent} from '../../services/ResultService/ResultService';
import { ISearchResults, ISearchResult } from '../../models/ISearchResult';
import IResultService from '../../services/ResultService/IResultService';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';
import TemplateService from '../../services/TemplateService/TemplateService';
import MockTemplateService from '../../services/TemplateService/MockTemplateService';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/Callout';
import { PropertyFieldLabelWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldLabelWithCallout';
import BingMap from '../../components/BingMap/BingMap';
import styles from '../../services/TemplateService/BaseTemplateService.module.scss';

export default class BingMapsSearchWebPart extends BaseClientSideWebPart<IBingMapsSearchWebPartProps> {

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
      
      //this.renderMockResults();

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

  private setMockResults() : void {
      
    let results : ISearchResult[] = [ {
          Title: "Title Item One",
          Description: "Description Item One",
          Category: "One,Two",
          Latitude: "13.0827",
          Longitude: "80.2707"
      }, {
          Title: "Title Item Two",
          Description: "Description Item Two",
          Category: "Two",
          Latitude: "14.0827",
          Longitude: "80.2707"
      }, {
          Title: "Title Item Three",
          Description: "Description Item Three", 
          Category: "One,Two,Three",
          Latitude: "13.0827",
          Longitude: "81.2707"
      }
    ];
    
    this._searchResults = {
      RelevantResults: results,
      QueryKeywords: "Test",
      RefinementResults:[]
    };      

  }

  public render(): void {
    
    if (Environment.type === EnvironmentType.Local) { this.setMockResults(); }

    this._map = React.createElement(
      BingMaps, {
        componentId: this.componentId, 
        pinResults: this._searchResults,
        templateService: this._templateService,
        bingMapsAPIKey: this.properties.bingMapsAPIKey,
        hbsTemplate: this.properties.hbsTemplate,
        inlineStyles: this.properties.inlineStyles,
        mapTypeId: this.properties.mapTypeId,
        zoom: this.properties.zoom,
        navigationBarMode: this.properties.navigationBarMode,
        supportedMapTypes: this.properties.supportedMapTypes,
        categoryIcons: this.properties.categoryIcons,
        columns: this.properties.columns,
        latitudeColumnName: this.properties.latitudeColumnName,
        longitudeColumnName: this.properties.longitudeColumnName,
        categoryColumnName: this.properties.categoryColumnName
      }
    );

    this._component = ReactDom.render(this._map, this.domElement) as BingMaps;
    
  }

  public onChangeHappened(e: ISearchEvent) {
    if(this._map) this._map.props.pinResults = e.results;
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
          return strings.requiredField;
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
                PropertyFieldNumber("zoom", {
                  key: "zoom",
                  label: strings.zoomLabel,
                  value: this.properties.zoom,
                  onGetErrorMessage: this._validateEmptyField.bind(this),
                  minValue: 1,
                  maxValue: 19
                }),
                PropertyPaneDropdown("mapTypeId", {
                  label: strings.mapTypeIdLabel,
                  selectedKey: this.properties.mapTypeId,
                  options: this._mapTypeIds
                }),
                PropertyFieldMultiSelect("supportedMapTypes", {
                  key: "supportedMapTypes",
                  label: strings.supportedMapTypesLabel,
                  options: this._mapTypeIds,
                  selectedKeys: this.properties.supportedMapTypes
                })
              ]
            },
            {
              groupName: "Column Configuration",
              groupFields: [
                PropertyPaneTextField('latitudeColumnName', {
                  label: strings.latitudeColumnNameLabel,
                  value: this.properties.latitudeColumnName,
                  onGetErrorMessage: this._validateEmptyField.bind(this)
                }),
                PropertyPaneTextField('longitudeColumnName', {
                  label: strings.longitudeColumnNameLabel,
                  value: this.properties.longitudeColumnName,
                  onGetErrorMessage: this._validateEmptyField.bind(this)
                }),                
                PropertyPaneTextField('categoryColumnName', {
                  label: strings.categoryColumnNameLabel,
                  value: this.properties.categoryColumnName
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
              groupName: strings.iconMappingsGroupLabel,
              groupFields: [
                PropertyFieldCollectionData('categoryIcons',{
                  key: 'categoryIcons',
                  label: strings.categoryIconsLabel,
                  panelHeader: strings.categoryIconsPanelHeader,
                  panelDescription: strings.categoryIconsPanelDesc,
                  manageBtnLabel: strings.categoryIconsButtonLabel,
                  value: this.properties.categoryIcons,
                  enableSorting: true,
                  fields: [{
                      id:"pattern",
                      title:strings.patternLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id:"url",
                      title: strings.urlLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id:"comparetype",
                      title: strings.compareTypeLabel,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "regex",
                          text: strings.compareTypeRegExText
                        },
                        {
                          key: "alltags",
                          text: strings.compareTypeAltTagText
                        }
                      ]
                    }
                  ]
                })
              ]
            },
            {
              groupName: strings.stylesTemplatesGroupLabel,
              groupFields: [                
                this._propertyFieldCodeEditor('inlineStyles', {
                    label: strings.inlineStylesTitle,
                    panelTitle: strings.inlineStylesPanelTitle,
                    initialValue: this.properties.inlineStyles,
                    deferredValidationTime: 500,
                    onPropertyChange: this.onPropertyPaneFieldChanged,
                    properties: this.properties,
                    key: 'inlineStylesCodeEditor',
                    language: this._propertyFieldCodeEditorLanguages.Handlebars
                }),                
                this._propertyFieldCodeEditor('hbsTemplate', {
                    label: strings.hbsTemplateLabel,
                    panelTitle: strings.hbsTemplatePanelTitle,
                    initialValue: this.properties.hbsTemplate,
                    deferredValidationTime: 500,
                    onPropertyChange: this.onPropertyPaneFieldChanged,
                    properties: this.properties,
                    key: 'hbsTemplateCodeEditor',
                    language: this._propertyFieldCodeEditorLanguages.Handlebars
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
