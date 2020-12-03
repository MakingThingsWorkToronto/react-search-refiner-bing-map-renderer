import * as React from 'react';
import IBingMapProps from './IBingMapProps';
import styles from './BingMap.module.scss';
import { ReactBingmaps } from 'react-bingmaps'; 
import { ISearchResult } from '../../models/ISearchResult';
import ICategoryIcon from '../../models/ICategoryIcon';

export default class BingMap extends React.Component<IBingMapProps, {}> {

    private _categoryIcons : ICategoryIcon[] = [];

    public render() {
        
        this._categoryIcons = this.props.categoryIcons ? this.props.categoryIcons.map((item) => {
            return {
                comparer : item.comparetype == "regex" ? this.createCompareRegEx(item.pattern) : this.createCompareList(item.pattern),
                url: item.url
            };
        }) : [];

        var results : ISearchResult[] = [];

        var r = (this.props.pinResults && this.props.pinResults.RelevantResults ? this.props.pinResults.RelevantResults : results);
        var resultPins = r.map(result => {
            return { 
                location: this.parseLocation(result[this.props.latitudeColumnName],result[this.props.longitudeColumnName]),
                addHandler: "mouseover",
                infoboxOption:  {
                    htmlContent: this.createMarkup(result)
                }, 
                pushPinOption: {
                    title: result.Title,
                    description: result.Description, 
                    icon: this.props.categoryColumnName ? this.getPinIcon(result[this.props.categoryColumnName]) : null
                }
            }; 
        }); 
        
        var validResultPins = [];
        resultPins.forEach((pin,index)=>{
            if(pin.location && !isNaN(pin.location[0]) && !isNaN(pin.location[1])) {
                validResultPins.push(pin);
            }
        });

        var bounds = validResultPins.map(pin => {
            return pin.location;  
        });

        var center = this.props.center;
        if(bounds.length == 1) { 
            center = bounds[0]; 
            bounds = [];
        }

        var p = (this.props.polygonResults && this.props.polygonResults.RelevantResults ? this.props.polygonResults.RelevantResults : results);
        var resultPolygons = p.map(result => {
            return {
                title: this.props.titleColumnName == "" ? "" : result[this.props.titleColumnName || "title"],
                fillColor: this.props.fillColor,
                strokeColor: this.props.strokeColor,
                strokeThickness: this.props.strokeThickness,
                target: result[this.props.targetColumnName || "target"],
                shape: result[this.props.polygonColumnName || "shape"] 
            };
        });
 
        var inlineStylesMapArea = {};
        var legendItems = [];
        var legendStyle = {
            display:"none"
        };

        if(this.props.height && this.props.height.trim().length>0) {
            inlineStylesMapArea = {
                height : this.props.height
            };
        }

        if(this.props.showLegend == true) {
            
            legendItems.push(<li><h3>Map Legend</h3></li>);

            var nextLegendItems =  this.props.categoryIcons.map((item,key)=>
                <li>
                    <img src={item.url} alt={item.legend} />
                    <span>{item.legend}</span>
                </li>
            );

            nextLegendItems.forEach((val,idx) => legendItems.push(val));
            
            legendStyle = {
                display: "block"
            };

        }

        return (
            <div className={styles.spReactBingMap} style={inlineStylesMapArea} >
                <div dangerouslySetInnerHTML={this.getInlineStyles()}></div>
                <ReactBingmaps
                    bingmapKey={this.props.bingMapsAPIKey} 
                    mapTypeId={this.props.mapTypeId}
                    navigationBarMode={this.props.navigationBarMode}
                    supportedMapTypes={this.props.supportedMapTypes}
                    infoboxesWithPushPins={validResultPins}
                    zoom={this.props.zoom}
                    className={styles.mapArea}
                    bounds={bounds}
                    center={center}
                    compressedPolygons={resultPolygons}
                    startOptions={this.props.startOptions}
                    mapOptions={this.props.mapOptions}
                    >
                    
                </ReactBingmaps>
                <ul className={styles.legendArea} style={legendStyle}> 
                    {legendItems}
                </ul>
            </div>
        );

    }

    private createCompareRegEx(pattern: string): any {

        return (fieldValue:string) => {
            try {
                var regex = new RegExp(pattern);
                return regex.test(fieldValue);
            } catch (e) {
                return false;
            }
        };
/*
        return (function(bits){
            return function(fieldValue:string) {
                try {
                    var regex = new RegExp(bits);
                    return regex.test(fieldValue);
                } catch (e) {
                    return false;
                }
            };
        })(pattern);
*/
    }

    private createCompareList(pattern:string) : any {

        return (fieldValue:string) => {
            try {

                var parts = pattern.split(","),
                    hasVal = true;

                parts.forEach((part,index)=>{
                    if(fieldValue.indexOf(part)==-1) {
                        hasVal = false;
                    }
                });
                
                return hasVal;

            } catch (e){

                return false;

            }
        };

        /*
        return (function(bits){

            return function(fieldValue:string) {
                
                try {

                    var parts = bits.split(","),
                        hasVal = true;

                    parts.forEach((part,index)=>{
                        if(fieldValue.indexOf(part)==-1) {
                            hasVal = false;
                        }
                    });
                    
                    return hasVal;

                } catch (e){

                    return false;

                }

            };

        })(pattern);
        */
    }


    
    private parseLocation(lat:string,long:string) : number[] {
        var loc = [lat,long],
            numCoords : number[] = loc.map(element => {
                try {
                    return parseFloat(element); 
                } catch(ex){
                    return 0;
                }
            });
        return numCoords;
    }

    private getInlineStyles(): any {
        return { __html: this.props.inlineStyles };
    }

    private createMarkup(result: ISearchResult) : string {

        let tmpl = this.props.hbsTemplate;
        try {
            return this.props.templateService.processTemplate({
                result: result,
                styles: styles
            }, tmpl);
        } catch (ex) {
            console.log("Failure generating template: " + ex.toString());
            return "";
        }

    }

    private getPinIcon(categoryValue: string) : string {
        if(!categoryValue) return null;
        let icon = null;
        let c = 0;
        for(c=0; c<this._categoryIcons.length; c++) {
            let mapping = this._categoryIcons[c];
            if(mapping.comparer(categoryValue)) {
                icon = mapping.url;
                break;
            }
        }
        return icon;
    }


}