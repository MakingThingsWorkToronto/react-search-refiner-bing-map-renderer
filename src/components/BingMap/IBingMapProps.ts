import { ISearchResults } from '../../models/ISearchResult';
import IResultService from '../../services/ResultService/IResultService';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';

export default interface IBingMapProps {
    componentId: string;
    pinResults?: ISearchResults;
    polygonResults?: ISearchResults;
    templateService: BaseTemplateService;
    bingMapsAPIKey: string;
    hbsTemplate?: string;
    inlineStyles?: string;
    mapTypeId: string;
    zoom: number;
    navigationBarMode: string;
    supportedMapTypes: string[];
    categoryIcons?: any[];
    columns: any[];
    center?: number[];
    latitudeColumnName?: string;
    longitudeColumnName?: string;
    categoryColumnName?: string;
    polygonColumnName?: string;
    titleColumnName?: string;
    targetColumnName?: string;
    fillColor? : string;
    strokeColor? : string;
    strokeThickness?: number;
}