import BaseTemplateService from                    './BaseTemplateService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

class TemplateService extends BaseTemplateService {

    private _spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient, locale: string) {

        super();
        this._spHttpClient = spHttpClient;
        this.CurrentLocale = locale;
    }

}

export default TemplateService;