import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown, IPropertyPaneDropdownOption, IPropertyPanePage, PropertyPaneToggle } from "@microsoft/sp-property-pane";

import * as strings from 'BingMapsSearchWebPartStrings';
import BingMaps from '../../components/BingMap/BingMap';
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

export default class BingMapsSearchWebPart extends BaseClientSideWebPart<IBingMapsSearchWebPartProps> {

  private _isInitialized: boolean = false;
  private _searchResults : ISearchResults;
  private _resultService: IResultService;
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

  protected async onInit(): Promise<void> {
    
    if (Environment.type === EnvironmentType.Local) {

      this._templateService = new MockTemplateService(this.context.pageContext.cultureInfo.currentUICultureName);
      
      //this.renderMockResults();

    } else {

        this._templateService = new TemplateService(this.context.spHttpClient, this.context.pageContext.cultureInfo.currentUICultureName);

    }

    await this._templateService.init();

    this._resultService = new ResultService();
    this._resultService.registerRenderer(
            this.componentId, 
            'Bing Maps', 
            'MapPin', 
            (e) => this.onChangeHappened(e), this.properties.columns.map((i) => { return i.name; }) as string[]
    );

    this._isInitialized = true;
    
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
      }, {
        Title: "Title Item Three",
        Description: "Description Item Three", 
        Category: "One,Two,Three",
        Latitude: "A",
        Longitude: "B"
      }
    ];
    
    this._searchResults = {
      RelevantResults: results,
      QueryKeywords: "Test",
      RefinementResults:[]
    };      

  }

  public render(): void {
    
    if(!this._isInitialized) return;

    if (Environment.type === EnvironmentType.Local) { this.setMockResults(); }

    const map = React.createElement(
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
        categoryColumnName: this.properties.categoryColumnName,
        height: this.properties.height,
        startOptions: {
          disableStreetside:this.properties.disableStreetside,
          showDashboard:this.properties.showDashboard,
          showLocateMeButton: this.properties.showLocateMeButton,
          showMapTypeSelector:this.properties.showMapTypeSelector,
          showScalebar:this.properties.showScalebar,
          showZoomButtons:this.properties.showZoomButtons
        },
        mapOptions: {
          disableZooming:this.properties.disableZooming,
          disableScrollWheelZoom:this.properties.disableScrollWheelZoom,
          allowInfoboxOverflow:this.properties.allowInfoboxOverflow,
          disableBirdseye:this.properties.disableBirdseye,
          disablePanning:this.properties.disablePanning,
          maxZoom: this.properties.maxZoom,
          minZoom: this.properties.minZoom
        },
        showLegend: this.properties.showLegend
      }
    );

    ReactDom.render(map, this.domElement);
    
  }

  public onChangeHappened(e: ISearchEvent) {
    this._searchResults = e.results;
    if(this._isInitialized) this.render();
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
          this.getOptionsPage(),
          this.getStylesPage(),
          this.getBingMapsOptionsPage()
      ]
    };
  }
  private getOptionsPage():IPropertyPanePage {
    return {
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
            PropertyPaneTextField('height', {
              label: strings.mapHeightLabel,
              value: this.properties.height
            }),
            PropertyFieldNumber("zoom", {
              key: "zoom",
              label: strings.zoomLabel,
              value: this.properties.zoom,
              onGetErrorMessage: this._validateEmptyField.bind(this),
              minValue: 1,
              maxValue: 19
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
        }
      ]
    };
  }

  private getStylesPage():IPropertyPanePage {
    return {
      header: {
        description:strings.StylesPaneDescription
      },
      groups: [
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
                  id:"legend",
                  title: strings.legendLabel,
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
            }),
            PropertyPaneToggle('showLegend', {
              label: strings.showLegendLabel, 
              checked: this.properties.showLegend
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
    };
  }

  
  private getBingMapsOptionsPage():IPropertyPanePage {
    return {
      header: {
        description:strings.BingMapsPageDescription
      },
      groups: [
        {
          groupName: strings.bingMapsGroupNameExtendedLabel,
          groupFields: [
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
            PropertyPaneToggle('disableZooming', {
              label: strings.disableZoomingLabel, 
              checked: this.properties.disableZooming
            }), 
            PropertyPaneToggle('disableScrollWheelZoom', {
              label: strings.disableScrollWheelZoomLabel, 
              checked: this.properties.disableScrollWheelZoom
            }), 
            PropertyPaneToggle('allowInfoboxOverflow', {
              label: strings.allowInfoboxOverflowLabel, 
              checked: this.properties.allowInfoboxOverflow
            }), 
            PropertyPaneToggle('disableBirdseye', {
              label: strings.disableBirdseyeLabel, 
              checked: this.properties.disableBirdseye
            }), 
            PropertyPaneToggle('disablePanning', {
              label: strings.disablePanningLabel, 
              checked: this.properties.disablePanning
            }), 
            PropertyPaneToggle('disableStreetside', {
              label: strings.disableStreetsideLabel, 
              checked: this.properties.disableStreetside
            }), 
            PropertyPaneToggle('showDashboard', {
              label: strings.showDashboardLabel, 
              checked: this.properties.showDashboard
            }), 
            PropertyPaneToggle('showLocateMeButton', {
              label: strings.showLocateMeButtonLabel, 
              checked: this.properties.showLocateMeButton
            }), 
            PropertyPaneToggle('showMapTypeSelector', {
              label: strings.showMapTypeSelectorLabel, 
              checked: this.properties.showMapTypeSelector
            }), 
            PropertyPaneToggle('showScalebar', {
              label: strings.showScalebarLabel, 
              checked: this.properties.showScalebar
            }), 
            PropertyPaneToggle('showZoomButtons', {
              label: strings.showZoomButtonsLabel, 
              checked: this.properties.showZoomButtons
            }), 
            PropertyFieldNumber("minZoom", {
              key: "minZoom",
              label: strings.minZoomLabel,
              value: this.properties.minZoom,
              minValue: 1,
              maxValue: 19
            }),
            PropertyFieldNumber("maxZoom", {
              key: "maxZoom",
              label: strings.maxZoomLabel,
              value: this.properties.maxZoom,
              minValue: 1,
              maxValue: 19
            })
          ]
        },
      ]
    };
  }

}
