
export default interface IBingMapsSearchWebPartProps {
    bingMapsAPIKey: string;
    hbsTemplate: string;
    mapTypeId: string;
    zoom: number;
    inlineStyles: string;
    navigationBarMode: string;
    supportedMapTypes: string[];
    categoryIcons: any[];
    columns: any[];
    latitudeColumnName: string;
    longitudeColumnName: string;
    categoryColumnName: string;
}
  