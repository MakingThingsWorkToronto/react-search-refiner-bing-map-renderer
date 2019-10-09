
export default interface IBingMapsSearchWebPartProps {
    bingMapsAPIKey: string;
    hbsTemplate: string;
    mapTypeId: string;
    height: string;
    zoom: number;
    inlineStyles: string;
    navigationBarMode: string;
    supportedMapTypes: string[];
    categoryIcons: any[];
    columns: any[];
    latitudeColumnName: string;
    longitudeColumnName: string;
    categoryColumnName: string;

    disableZooming:boolean;
    disableScrollWheelZoom:boolean;
    allowInfoboxOverflow:boolean;
    disableBirdseye:boolean;
    disablePanning:boolean;

    disableStreetside:boolean;
    showDashboard:boolean;
    showLocateMeButton: boolean;
    showMapTypeSelector:boolean;
    showScalebar:boolean;
    showZoomButtons:boolean;
    
    minZoom: number;
    maxZoom: number;

    showLegend: boolean;
    
}
  