import 'core-js/modules/es7.array.includes.js';
import 'core-js/modules/es6.string.includes.js';
import 'core-js/modules/es6.number.is-nan.js';
import * as Handlebars from 'handlebars';
import { ISearchResult } from '../../models/ISearchResult';
import { isEmpty, uniqBy, uniq } from '@microsoft/sp-lodash-subset';
import * as strings from 'BingMapsSearchWebPartStrings';
import { Text } from '@microsoft/sp-core-library';
import * as HandlebarsGroupBy from 'handlebars-group-by';

abstract class BaseTemplateService {

    private _handleBarsInstance : typeof Handlebars;

    public CurrentLocale : string = "en";

    constructor() {
        
        // Create local instance of handlebars
        this._handleBarsInstance = Handlebars.create();        
        
        // Registers all helpers
        this.registerTemplateServices();

    }

    public async init() : Promise<void> {

        // Registers handlerbars-helpers
        await this.LoadHandlebarsHelpers();

    }

    private async LoadHandlebarsHelpers() {
        if ((<any>window).mapsHBHelper !== undefined) {
            // early check - seems to never hit(?)
            return;
        }
        let component = await import(
            /* webpackChunkName: 'search-handlebars-helpers' */
            'handlebars-helpers'
        );
        (<any>window).mapsHBHelper = component({
            handlebars: this._handleBarsInstance
        });
    }

    /**
     * Registers useful helpers for search results templates
     */
    private registerTemplateServices() {

        // Register the group by helper
        HandlebarsGroupBy.register(this._handleBarsInstance);

        // Return the URL of the search result item
        // Usage: <a href="{{url item}}">
        this._handleBarsInstance.registerHelper("getUrl", (item: ISearchResult) => {
            if (!isEmpty(item))
                return item.ServerRedirectedURL ? item.ServerRedirectedURL : item.Path;
        });

        // Return the search result count message
        // Usage: {{getCountMessage totalRows keywords}} or {{getCountMessage totalRows null}}
        this._handleBarsInstance.registerHelper("getCountMessage", (totalRows: string, inputQuery?: string) => {

            const countResultMessage = inputQuery ? Text.format(strings.CountMessageLong, totalRows, inputQuery) : Text.format(strings.CountMessageShort, totalRows);
            return new Handlebars.SafeString(countResultMessage);
        });

        // Return the preview image URL for the search result item
        // Usage: <img src="{{previewSrc item}}""/>
        this._handleBarsInstance.registerHelper("getPreviewSrc", (item: ISearchResult) => {

            let previewSrc = "";

            if (item) {
                if (!isEmpty(item.SiteLogo)) previewSrc = item.SiteLogo;
                else if (!isEmpty(item.PreviewUrl)) previewSrc = item.PreviewUrl;
                else if (!isEmpty(item.PictureThumbnailURL)) previewSrc = item.PictureThumbnailURL;
                else if (!isEmpty(item.ServerRedirectedPreviewURL)) previewSrc = item.ServerRedirectedPreviewURL;
            }

            return previewSrc;
        });

        // Return the highlighted summary of the search result item
        // <p>{{summary HitHighlightedSummary}}</p>
        this._handleBarsInstance.registerHelper("getSummary", (hitHighlightedSummary: string) => {
            if (!isEmpty(hitHighlightedSummary)) {
                return new Handlebars.SafeString(hitHighlightedSummary.replace(/<c0\>/g, "<strong>").replace(/<\/c0\>/g, "</strong>").replace(/<ddd\/>/g, "&#8230;"));
            }
        });

        // Return the formatted date according to current locale using moment.js
        // <p>{{getDate Created "LL"}}</p>
        this._handleBarsInstance.registerHelper("getDate", (date: string, format: string) => {
            try {
                let d = (<any>window).mapsHBHelper.moment(date, format, { lang: this.CurrentLocale, datejs: false });
                return d;
            } catch (error) {
                return date;
            }
        });

        // Return the URL or Title part of a URL automatic managed property
        // <p>{{getUrlField MyLinkOWSURLH "Title"}}</p>
        this._handleBarsInstance.registerHelper("getUrlField", (urlField: string, value: "URL" | "Title") => {
            if (!isEmpty(urlField)) {
                let separatorPos = urlField.indexOf(",");
                if (separatorPos === -1) {
                    return urlField;
                }
                if (value === "URL") {
                    return urlField.substr(0, separatorPos);
                }
                return urlField.substr(separatorPos + 1).trim();
            }
            return urlField;
        });

        // Return the unique count based on an array or property of an object in the array
        // <p>{{getUniqueCount items "Title"}}</p>
        this._handleBarsInstance.registerHelper("getUniqueCount", (array: any[], property: string) => {
            if (!Array.isArray(array)) return 0;
            if (array.length === 0) return 0;

            let result;
            if (property) {
                result = uniqBy(array, property);

            }
            else {
                result = uniq(array);
            }
            return result.length;
        });
    }

    /**
     * Compile the specified Handlebars template with the associated context object¸
     * @returns the compiled HTML template string 
     */
    public processTemplate(templateContext: any, templateContent: string): string {
        
        let template = this._handleBarsInstance.compile(templateContent);
        let result = template(templateContext);

        return result;
    }

}

export default BaseTemplateService;