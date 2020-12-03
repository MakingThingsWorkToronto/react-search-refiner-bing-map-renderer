import BaseTemplateService from                    './BaseTemplateService';

class MockTemplateService extends BaseTemplateService {

    constructor(locale: string) {
        super();    
        this.CurrentLocale = locale;
    }

}

export default MockTemplateService;