import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class EntityReference {
    id: string;
    typeName: string;
    constructor(typeName: string, id: string) {
        this.id = id;
        this.typeName = typeName;
    }
    static get(e: ComponentFramework.WebApi.Entity, name: string): EntityReference {
        return new EntityReference(
            e[`_${name}_value@Microsoft.Dynamics.CRM.lookuplogicalname`],
            e[`_${name}_value`]
        );
    }
}

export class PrimaryEntity {
    Entity: EntityReference;
    constructor(context: ComponentFramework.Context<IInputs>) {
        this.Entity = new EntityReference(
            (context as any).page.entityTypeName,
            (context as any).page.entityId
        );
    }
}

export class FetchXML {
    static prepareOptions(fetchXml: string): string {
        return `?fetchXml=${encodeURIComponent(fetchXml)}`;
    }
}

export const isInHarness = (): boolean => {
    return window.location.href.indexOf(".dynamics.com") < 0;
}

export class SharePointHelper {
    private spList: string[];
    private apiUrl: string;
    constructor(site: string, apiUrl: string) {
        this.spList = site.split(',');
        this.apiUrl = apiUrl;
    }
    public makeApiUrl(url: string): string {
        let spFilePath = "", spSiteUrl: string = "";
        for (let i = 0; i < this.spList.length; i++) {
            if (url.indexOf(this.spList[i]) > -1) {
                spSiteUrl = encodeURIComponent(this.spList[i]);
                spFilePath = encodeURIComponent(url.replace(this.spList[i], ""));
                break;
            }
        }
        return this.apiUrl.replace("{0}", spSiteUrl)
            .replace("{1}", spFilePath);
    }
}