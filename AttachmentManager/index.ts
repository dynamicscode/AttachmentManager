import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { AttachmentManagerApp } from "./AttachmentManagerApp";
import { colGroupProperties } from "@uifabric/utilities";
import { RGBA_REGEX, CollapseAllVisibility } from "office-ui-fabric-react";

class EntityReference {
	id: string;
	typeName: string;
	constructor(typeName: string, id: string) {
		this.id = id;
		this.typeName = typeName;
	}
}

export class AttachmentManager implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	private async getEmail(id: string) {
		this._context.webAPI.retrieveRecord("email", id).then(this.renderControl);
		const email = await this._context.webAPI.retrieveRecord("email", id);
		return email;
	}

	private renderControl(): void{
		console.log("renderControl");
		ReactDOM.render(
			React.createElement(AttachmentManagerApp)
			, this._container
		);
	}

	private async getSharePointDocuments(id: string, entityName: string) {
		const fetchXml: string = `
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
				<link-entity name="${entityName}" from="${entityName}id" to="regardingobjectid" alias="bb">
					<filter type="and">
						<condition attribute="${entityName}id" operator="eq" uitype="${entityName}id" value="${id}"/>
					</filter>
				</link-entity>
			</entity>
		</fetch>`;
		const query = `?fetchXml=${encodeURIComponent(fetchXml)}`;
		const documents = await this._context.webAPI.retrieveMultipleRecords("sharepointdocument", query);
		return documents.entities;
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this._context = context;
		this._container = container;

		const regarding: EntityReference = new EntityReference(
			(<any>this._context).page.entityTypeName,
			(<any>this._context).page.entityId
		);

		this.getEmail(regarding.id).then(
			(e) => {
				const regardingEntityName:string = e["_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
				const regardingObjectId: string = e["_regardingobjectid_value"];
		
				this.getSharePointDocuments(regardingObjectId, regardingEntityName).then(
					(ec) => {
						console.log(`No. of documents in SP ${ec.length}`);

						this.renderControl();
					}
				);
			}
		)
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		ReactDOM.unmountComponentAtNode(this._container);
	}
}