import { IInputs } from "./generated/ManifestTypes";
import { FetchXML, EntityReference } from "./PCFHelper";

export const ActivityMimeAttachment = {
    EntityName: "activitymimeattachment",
    create: async(content: any, regarding: EntityReference, name: string, context: ComponentFramework.Context<IInputs>) => {
        
        var attachment = {} as any;
        attachment.body = content;
        attachment["objectid_activitypointer@odata.bind"] = `activitypointers(${regarding.id})`;
        attachment["objecttypecode"] = regarding.typeName;
        attachment["filename"] = name;

        await context.webAPI.createRecord(ActivityMimeAttachment.EntityName, attachment);
    }
}

export const Email = {
    EntityName: "email",
    RegardingObject: "regardingobjectid",
    getById: async(id: string, context: ComponentFramework.Context<IInputs>) => {
        const email = await context.webAPI.retrieveRecord(Email.EntityName, id);
        return email;
    }
}

export const SharePointDocument = {
    EntityName: "sharepointdocument",
    FetchXml: `
		<fetch distinct="false" mapping="logical" returntotalrecordcount="true" page="1" count="100" no-lock="false">
			<entity name="sharepointdocument">
				<attribute name="documentid"/>
				<attribute name="fullname"/>
				<attribute name="relativelocation"/>
				<attribute name="sharepointcreatedon"/>
				<attribute name="ischeckedout"/>
				<attribute name="filetype"/>
				<attribute name="modified"/>
				<attribute name="sharepointmodifiedby"/>
				<attribute name="servicetype"/>
				<attribute name="absoluteurl"/>
				<attribute name="title"/>
				<attribute name="author"/>
				<attribute name="sharepointdocumentid"/>
				<attribute name="readurl"/>
				<attribute name="editurl"/>
				<attribute name="locationid"/>
				<attribute name="iconclassname"/>
				<attribute name="locationname"/>
				<order attribute="relativelocation" descending="false"/>
				<filter>
					<condition attribute="isrecursivefetch" operator="eq" value="0"/>
				</filter>
				<link-entity name="{entityName}" from="{entityName}id" to="regardingobjectid" alias="bb">
					<filter type="and">
						<condition attribute="{entityName}id" operator="eq" uitype="{entityName}id" value="{id}"/>
					</filter>
				</link-entity>
			</entity>
        </fetch>`,
        FullName : "fullname",
        AbsoluteUrl : "absoluteurl",
        FileType: "filetype",
        IconClassName: "iconclassname",
        LastModifiedOn: "modified",
        LastModifiedBy: "sharepointmodifiedby",
        getByRegarding: async(id: string, entityName: string, context: ComponentFramework.Context<IInputs>) => {
            const fetchXml: string = (SharePointDocument.FetchXml as string).split("{entityName}").join(entityName).split("{id}").join(id);

            const documents = await context.webAPI.retrieveMultipleRecords(SharePointDocument.EntityName, FetchXML.prepareOptions(fetchXml));
            return documents.entities;
        }

}
