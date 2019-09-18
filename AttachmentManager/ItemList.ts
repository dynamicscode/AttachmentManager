import { IColumn, IColor, ThemeSettingName } from "office-ui-fabric-react";
import { getLorem } from "./Lorem";


export interface IFileItem {
    key: number | string;
    id: string;
    fileName: string;
    fileType: string;
    fileUrl: string;
    lastModifiedOn: Date;
    lastModifiedBy: string;
    iconclassname: string;
}

export class ItemList {
    private columns: IColumn[];
    private items: IFileItem[];

    constructor() {
        this.columns = [];
        this.items = [];

        this.setColumns();
    }

    private setColumns(): void {
        this.columns = [];
        this.columns.push({
            key: 'icon',
            name: '',
            fieldName: 'iconclassname',
            minWidth: 20,
            maxWidth: 40,
            isResizable: false
        });
        this.columns.push({
            key: 'fileName',
            name: 'Name',
            fieldName: 'fileName',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
        this.columns.push({
            key: 'fileType',
            name: 'Type',
            fieldName: 'fileType',
            minWidth: 50,
            maxWidth: 50,
            isResizable: true
        });
        this.columns.push({
            key: 'lastModifiedOn',
            name: 'Last Modified On',
            fieldName: 'lastModifiedOn',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
        this.columns.push({
            key: 'lastModifiedBy',
            name: 'Last Modified By',
            fieldName: 'lastModifiedBy',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
    }

    private addMockItems(): void {
        this.items = [];
        for (let i = 1; i < 31; i++) {
            this.items.push({
                key: i,
                id: i.toString(),
                fileName: getLorem(4),
                fileType: getLorem(4),
                fileUrl: getLorem(4),
                iconclassname: getLorem(4),
                lastModifiedOn: new Date(),
                lastModifiedBy: getLorem(4)
            });
        }
    }

    public getColumns(): IColumn[] {
        return this.columns;
    }

    public getItems(): IFileItem[] {
        return this.items;
    }

    public setItems(items: IFileItem[]): void {
        if (items)
            this.items = items;
        else 
            this.addMockItems();
    }
}