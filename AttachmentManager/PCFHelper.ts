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