import { controlListBuilder, ControlListBuilderParams, GenericControlFactory, InfoBox } from "itdevcontrols";
import { Mapping, AjaxTransport, LocalListTransport, AjaxListTransport } from 'itdevtransports';
import { SPPageContext} from 'itdevcontexts';

// configuration JSON
export class FollowedSitesConfig {
    public limit: number;
    public viewAllLabelText: string;
}

export class FollowedSites {
    
    constructor(config: FollowedSitesConfig) {
        let limit = config.limit;
        let viewAllLabelText = config.viewAllLabelText;
        let listTemplate: string = `<div></div>`;
        
        let transport = new AjaxTransport({
            url: 'https://kksandbox.sharepoint.com/sites/aa/_api/social.following/my/followed(types=4)',
            // types: documents=2, sites=4, tags=8, all=2+4+8=14
            responseType: 'json',
            additionalHeaders: [
                { requestHeader: 'Content-Type', value: 'application/json; charset=UTF-8' }
            ],
        });

        // Url: "*/_layouts/15/sharepoint.aspx?v=following"
        function createAbsoluteFollowingUrl(): Promise<string> {
            
            return SPPageContext.getInstance().then(function(response){
                let webAbsoluteUrl = response.webAbsoluteUrl;
                let siteServerRelativeUrl = response.siteServerRelativeUrl;
                let suffix = "/_layouts/15/sharepoint.aspx?v=following";
                let absoluteFollowingUrl = webAbsoluteUrl.replace(siteServerRelativeUrl,"") + suffix;
                return absoluteFollowingUrl;
            })
        }

        transport.read().then((response) => {
            createAbsoluteFollowingUrl().then((url) => {
                response.data.value.reverse();
                let followedSitesTransport = new LocalListTransport({
                    data: response.data.value,
                    rowLimit: limit
                });
                let viewAllTransport = new LocalListTransport({
                    data: [{
                        Uri: url,
                        Name: viewAllLabelText,
                        Class: "it-dev-followed-sites-last-row"
                    }]
                });

                let joinedTransport = followedSitesTransport.union(viewAllTransport);

                // (done) 1. Odwracamy liste, obcinamy rzeczy
                // (done) 2. Omawiamy z BSz czy zrobić PULL REQUEST z reversem do transportu
                // (done) 3. Parametryzujemy link view all za pomocą _spPageContextInfo

                let listControlMapping: Mapping[] = [
                {
                    source: null,
                    target: 'template',
                    targetDefault: `
                    <div :class='dataModel.Class'>
                        <div class='followed-item'><a :href="dataModel.Uri">{{ dataModel.Name }}</a></div>
                    </div>
                `
                }];
                
                // Container implemented in a script editor an a publishing site 
                // <div id='FollowedSites'></div>
                let container = "#FollowedSites"; 
                
                let params = new ControlListBuilderParams<any>({
                    scope: container,
                    template: listTemplate,
                    controlFactory: new GenericControlFactory(InfoBox, listControlMapping),
                    transport: joinedTransport
                });
                controlListBuilder(params);
            });
        });
        
    }
}