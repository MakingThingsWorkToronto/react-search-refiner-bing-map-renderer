import { ISearchResults } from "../../models/ISearchResult";
import IResultService from "./IResultService";
import 'custom-event-polyfill';
import { render } from "react-dom";

export interface ISearchEvent extends CustomEvent {
    rendererId?: string;
    results?: ISearchResults;
    mountNode?: string;
    customTemplateFieldValues?: ICustomTemplateFieldValue[];
}

export interface IRenderer {
    id: string;
    name: string;
    icon: string;
    customFields?: string[];
}

export interface ICustomTemplateFieldValue {
    fieldName: string;
    searchProperty: string;
}

export class ResultService implements IResultService {
    private SEARCH_CHANGED_EVENT_NAME: string = "pnp-spfx-search-changed";
    private SEARCH_RENDERERS_OBJECT_NAME: string = "pnp-spfx-search-renderers";
    private _results: ISearchResults;
    private _renderEvent = null;
    public get results(): ISearchResults { return this._results; }

    private _isLoading: boolean;
    public get isLoading(): boolean { return this._isLoading; }
    public set isLoading(status: boolean) { this._isLoading = status; }

    public updateResultData(results: ISearchResults, rendererId: string, mountNode: string, customTemplateFieldValues?: ICustomTemplateFieldValue[]) {
        console.log("Data updated: " + rendererId);
        this._results = results;
        let searchEvent: ISearchEvent = new CustomEvent(this.SEARCH_CHANGED_EVENT_NAME);
        searchEvent.rendererId = rendererId;
        searchEvent.results = results; 
        searchEvent.mountNode = mountNode;
        searchEvent.customTemplateFieldValues = customTemplateFieldValues;
        window.dispatchEvent(searchEvent);
    }

    public registerRenderer(rendererId: string, rendererName: string, rendererIcon: string, callback: (e: ISearchEvent) => void, customFields?: string[]): void {
        const newRenderer = {
            id: rendererId,
            name: rendererName,
            icon: rendererIcon,
            customFields: customFields
        };

        this._renderEvent = this.handleNewDataRegistered.bind(this, rendererId, callback);
        
        if(window[this.SEARCH_RENDERERS_OBJECT_NAME] === undefined) window[this.SEARCH_RENDERERS_OBJECT_NAME] = [];

        const alreadyRegistered = window[this.SEARCH_RENDERERS_OBJECT_NAME].some((renderer)=>{ return renderer.id === rendererId; });

        if(alreadyRegistered) this.unregisterRenderer(rendererId);

        window[this.SEARCH_RENDERERS_OBJECT_NAME].push(newRenderer);      

        addEventListener(this.SEARCH_CHANGED_EVENT_NAME, this._renderEvent);

    }

    public unregisterRenderer(rendererId:string) {

        if(window[this.SEARCH_RENDERERS_OBJECT_NAME] !== undefined) {
            
            window[this.SEARCH_RENDERERS_OBJECT_NAME] = window[this.SEARCH_RENDERERS_OBJECT_NAME].filter((renderer)=>{
                return renderer.id !== rendererId;
            });
            
            removeEventListener(this.SEARCH_CHANGED_EVENT_NAME, this._renderEvent);

        }

    }

    public getRegisteredRenderers(): IRenderer[] {
        return window[this.SEARCH_RENDERERS_OBJECT_NAME];
    }

    private handleNewDataRegistered(rendererId, callback: (e) => void, e: ISearchEvent) {
        console.log("Handle new data registered: " + e.rendererId);
        console.log("Waiting for: " + rendererId);
        console.log(arguments);
        if(e.rendererId === rendererId) {
            console.log("Executing callback");
            if(window.location.href.indexOf("Mode=Edit") === -1){
                const searchWpId = e.mountNode;
                const eventElement = document.querySelector("#" + searchWpId).parentElement.parentElement;
                const wpSection = eventElement.closest(".CanvasZone");
                if(wpSection) {
                    console.log("Hiding source web part zone");
                    wpSection["style"].display = "none";
                }
            }
            callback(e);
        }
    }
}