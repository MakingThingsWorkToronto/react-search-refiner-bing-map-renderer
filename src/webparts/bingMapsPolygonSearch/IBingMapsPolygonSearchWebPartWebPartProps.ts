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
}
  