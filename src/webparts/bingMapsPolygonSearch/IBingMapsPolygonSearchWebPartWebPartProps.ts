export default interface IBingMapsPolygonSearchWebPartWebPartProps {
    bingMapsAPIKey: string;
    mapTypeId: string;
    zoom: number;
    navigationBarMode: string;
    supportedMapTypes: string[];
    columns: any[];
    titleColumnName: string;
    polygonColumnName: string;
    targetColumnName: string;
    fillColor: string;
    strokeColor: string;
    strokeThickness: number;
    centerLatitude: string;
    centerLongitude: string;
    showLabels: boolean;
    height:string;

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

    minZoom:number;
    maxZoom:number;
    
}
  