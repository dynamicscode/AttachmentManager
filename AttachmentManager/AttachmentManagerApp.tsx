import * as React from 'react';

import { DefaultButton, Stack, ProgressIndicator } from 'office-ui-fabric-react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import {
    DetailsList,
    DetailsListLayoutMode,
    IDetailsHeaderProps,
    Selection,
    IColumn,
    ConstrainMode,
    IDetailsFooterProps,
    DetailsRow
} from 'office-ui-fabric-react/lib/DetailsList';
import { IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { SelectionMode } from 'office-ui-fabric-react/lib/Selection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';

export interface IAttachmentProps {
    regardingObjectId: string;
    regardingEntityName: string;
    files: IFileItem[];
    onAttach: (selectedFiles: IFileItem[]) => Promise<void>;
}

export interface IAttachmentState {
    files: IFileItem[];
    hiddenModal: boolean;
    isInProgress: boolean;
}

export interface IFileItem {
    key: number | string;
    id: string;
    fileName: string;
    fileType: string;
    fileUrl: string;
    lastModifiedOn?: string;
    lastModifiedBy?: string;
    iconclassname: string;
}

// const _footerItem: IFileItem = {
//     key: 'Key',
//     id: 'Id',
//     fileName: 'Name',
//     fileType: 'Type',
//     fileUrl: '',
//     lastModifiedOn: 'Last Modified On',
//     iconclassname: '',
//     lastModifiedBy: 'Last Modified By'
// };

const classNames = mergeStyleSets({
    wrapper: {
        height: '80vh',
        position: 'relative'
    },
    filter: {
        paddingBottom: 20,
        maxWidth: 300
    },
    header: {
        margin: 0
    },
    row: {
        display: 'inline-block'
    }
});

const LOREM_IPSUM = (
    'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut ' +
    'labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut ' +
    'aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
    'eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt '
).split(' ');

let loremIndex = 0;

function _lorem(wordCount: number): string {
    const startIndex = loremIndex + wordCount > LOREM_IPSUM.length ? 0 : loremIndex;
    loremIndex = startIndex + wordCount;
    return LOREM_IPSUM.slice(startIndex, loremIndex).join(' ');
}

export class AttachmentManagerApp extends React.Component<IAttachmentProps, IAttachmentState> {
    private _selection: Selection;
    private _allFiles: IFileItem[];
    private _columns: IColumn[];

    constructor(props: IAttachmentProps) {
        super(props)

        initializeIcons();

        this._addItems();
        this._setColumns();

        this._selection = new Selection();

        if (this.props.files) {
            this._allFiles = this.props.files;
         } else { 
             this._addItems();
        };
        
        this.state = {
            files: this._allFiles,
            hiddenModal: true,
            isInProgress: false
        };

        this.attachFilesClicked = this.attachFilesClicked.bind(this);
        this.onFilterChanged = this.onFilterChanged.bind(this);
        this.onAttachClicked = this.onAttachClicked.bind(this);
        this.hideDialog = this.hideDialog.bind(this);
    }

    private _setColumns(): void {
        this._columns = [];
        this._columns.push({
            key: 'icon',
            name: '',
            fieldName: 'iconclassname',
            minWidth: 20,
            maxWidth: 40,
            isResizable: false
        });
        this._columns.push({
            key: 'fileName',
            name: 'File Name',
            fieldName: 'fileName',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
        this._columns.push({
            key: 'fileType',
            name: 'File Type',
            fieldName: 'fileType',
            minWidth: 50,
            maxWidth: 50,
            isResizable: true
        });
        this._columns.push({
            key: 'lastModifiedOn',
            name: 'Last Modified On',
            fieldName: 'lastModifiedOn',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
        this._columns.push({
            key: 'lastModifiedBy',
            name: 'Last Modified By',
            fieldName: 'lastModifiedBy',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
    }

    private _addItems(): void {
        this._allFiles = [];
        for (let i = 1; i < 31; i++) {
            this._allFiles.push({
                key: i,
                id: i.toString(),
                fileName: _lorem(4),
                fileType: _lorem(4),
                fileUrl: _lorem(4),
                iconclassname: _lorem(4)
            });
        }
    }

    public render(): JSX.Element {
        const { hiddenModal: hiddenDialog, files } = this.state;
        return (
            <div>
                <CommandBar
                    items={this.getItems()}
                />
                <Dialog
                    hidden={hiddenDialog}
                    onDismiss={() => { this.setState({ hiddenModal: true }) }}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Attach files',
                        subText: 'Choose files you want to attach to the email'
                    }}
                    modalProps={{
                        isBlocking: false,
                    }}
                    minWidth='800px'
                >
                    <div style={{ 'height': '400px' }}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <Sticky stickyPosition={StickyPositionType.Header}>
                                <Stack horizontal tokens={{childrenGap: 20, padding:10}}>
                                    <Stack.Item>
                                        <DefaultButton text="Attach" onClick={this.onAttachClicked} />
                                    </Stack.Item>
                                    <Stack.Item grow align="stretch">
                                        <SearchBox styles={{ root: { width: '100%' } }} placeholder="Search file" onChange={this.onFilterChanged} />
                                    </Stack.Item>
                                </Stack>
                                <Stack>
                                    { this.state.isInProgress && <ProgressIndicator label="In progress" description="Copying files from SharePoint to an email" /> }
                                </Stack>
                            </Sticky>
                            <MarqueeSelection selection={this._selection}>
                                <DetailsList
                                    items={files}
                                    columns={this._columns}
                                    setKey="set"
                                    layoutMode={DetailsListLayoutMode.fixedColumns}
                                    constrainMode={ConstrainMode.unconstrained}
                                    onRenderDetailsHeader={onRenderDetailsHeader}
                                    //onRenderDetailsFooter={onRenderDetailsFooter}
                                    selection={this._selection}
                                    selectionPreservedOnEmptyClick={true}
                                    ariaLabelForSelectionColumn="Toggle selection"
                                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                    onItemInvoked={this.onItemInvoked}
                                />
                            </MarqueeSelection>
                        </ScrollablePane>
                    </div>
                </Dialog>
            </div>
        );
    }

    private getItems = () => {
        return [
            {
                key: 'attachFile',
                name: 'Attach Files',
                cacheKey: 'myCacheKey',
                iconProps: {
                    iconName: 'Attach'
                },
                ariaLabel: 'Attach Files',
                onClick: this.attachFilesClicked
            }
        ]
    }

    private attachFilesClicked(): void {
        this.setState({ hiddenModal: false, isInProgress : false });
    }

    private onAttachClicked(): void {
        this.setState({isInProgress : true});
        this.props.onAttach(this.getSelectedFiles()).then(this.hideDialog);
    }

    private hideDialog(): void {
        this.setState({hiddenModal:true});
    }

    private onItemInvoked(item: IFileItem): void {
        //alert('Item invoked: ' + item.fileName);
    }

    private onFilterChanged(ev?: React.ChangeEvent<HTMLInputElement>, text?: string): void {
        this.setState({
            files: text ? this._allFiles.filter((item: IFileItem) => 
            hasText(item, text)) : this._allFiles
        });
    };

    private getSelectedFiles(): IFileItem[] {
        let selectedFileId: IFileItem[];
        selectedFileId = [];
        var fileUrl = '';

        for(let i = 0; i < this._selection.getSelectedCount(); i++) {
            selectedFileId.push((this._selection.getSelection()[i] as IFileItem));
        }

        return selectedFileId;
    }
}

function hasText(item: IFileItem, text: string): boolean {
    return `${item.id}|${item.fileName}|${item.fileType}`.indexOf(text) > -1;
}

function onRenderDetailsHeader(
    props?: IDetailsHeaderProps,
    defaultRender?: IRenderFunction<IDetailsHeaderProps>
): JSX.Element {
    return (
        <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
            {defaultRender && defaultRender({ ...props! })}
        </Sticky>
    );
}

// function onRenderDetailsFooter(props?: IDetailsFooterProps, defaultRender?: IRenderFunction<IDetailsFooterProps>): JSX.Element {
//     return (
//         <Sticky stickyPosition={StickyPositionType.Footer} isScrollSynced={true}>
//             <div className={classNames.row}>
//                 <DetailsRow
//                     columns={props!.columns}
//                     item={_footerItem}
//                     itemIndex={-1}
//                     selection={props!.selection}
//                     selectionMode={(props!.selection && props!.selection.mode) || SelectionMode.none}
//                     viewport={props!.viewport}
//                 />
//             </div>
//         </Sticky>
//     );
// }