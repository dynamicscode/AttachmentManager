interface IIconMap {
	SharePointIconClass: string;
	FabricUIIconClass: string;
}

export class IconMapper {
	private iconMappers: IIconMap[];
	constructor() {
		this.iconMappers = [];

		this.iconMappers.push({ SharePointIconClass: "excelFile", FabricUIIconClass : "ExcelDocument"});
		this.iconMappers.push({ SharePointIconClass: "pdfFile", FabricUIIconClass : "PDF"});
		this.iconMappers.push({ SharePointIconClass: "powerPointFile", FabricUIIconClass : "PowerPointDocument"});
		this.iconMappers.push({ SharePointIconClass: "text", FabricUIIconClass : "TextDocument"});
		this.iconMappers.push({ SharePointIconClass: "imageFile", FabricUIIconClass : "FileImage"});
		this.iconMappers.push({ SharePointIconClass: "wordFile", FabricUIIconClass : "WordDocument"});
	}

	public getBySharePointIcon(icon: string): string {
		const r = this.iconMappers.filter(v => icon.indexOf(v.SharePointIconClass) > -1);
		if (r.length > 0)
			return r[0].FabricUIIconClass;
		return icon;
	}
}