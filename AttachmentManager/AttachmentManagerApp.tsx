import * as React from 'react';

import { DefaultButton, PrimaryButton, Stack, IStackTokens, CommandBarButton } from 'office-ui-fabric-react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
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
    fileLists: IFileItem[];
    onAttach: (selectedFiles: IFileItem[]) => void;
}

export interface IAttachmentState {
    fileLists: IFileItem[];
    hiddenModal: boolean;
    selectedFiles: IFileItem[];
}

export interface IFileItem {
    id: string;
    fileName: string;
    fileType: string;
    fileUrl: string;
    iconclassname: string;
}

const _footerItem: IFileItem = {
    id: 'Id',
    fileName: 'Name',
    fileType: 'Type',
    fileUrl: '',
    iconclassname: ''
};

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

        this._selection = new Selection({
            onSelectionChanged: () => this.setState({ selectedFiles: this.getSelectedFiles() })
        });

        this.state = {
            fileLists: this.props.fileLists ? this.props.fileLists : this._allFiles, 
            hiddenModal: true,
            selectedFiles: this.getSelectedFiles()
        };

        this.attachFilesClicked = this.attachFilesClicked.bind(this);
        this.onFilterChanged = this.onFilterChanged.bind(this);
        this.onAttachClicked = this.onAttachClicked.bind(this);
    }

    private _setColumns(): void {
        this._columns = [];
        this._columns.push({
            key: 'id',
            name: 'Id',
            fieldName: 'id',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
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
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        });
    }

    private _addItems(): void {
        this._allFiles = [];
        for (let i = 1; i < 31; i++) {
            this._allFiles.push({
                id: i.toString(),
                fileName: _lorem(4),
                fileType: _lorem(4),
                fileUrl: _lorem(4),
                iconclassname: _lorem(4)
            });
        }
    }

    public render(): JSX.Element {
        const { hiddenModal: hiddenDialog, fileLists, selectedFiles } = this.state;
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
                    <div style={{ 'height': '600px' }}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <Sticky stickyPosition={StickyPositionType.Header}>
                                <Stack horizontal>
                                    <TextField styles={{ root: { width: 250 } }} placeholder="Search file" onChange={this.onFilterChanged} />
                                    <DefaultButton text="Attach" onClick={this.onAttachClicked} />
                                </Stack> 
                            </Sticky>
                            <MarqueeSelection selection={this._selection}>
                                <DetailsList
                                    items={fileLists}
                                    columns={this._columns}
                                    setKey="set"
                                    layoutMode={DetailsListLayoutMode.fixedColumns}
                                    constrainMode={ConstrainMode.unconstrained}
                                    onRenderDetailsHeader={onRenderDetailsHeader}
                                    onRenderDetailsFooter={onRenderDetailsFooter}
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
        this.setState({ hiddenModal: false });
    }

    private onAttachClicked(): void {
        this.props.onAttach(this.state.selectedFiles);
        this.setState({ hiddenModal: true });
    }

    private onItemInvoked(item: IFileItem): void {
        alert('Item invoked: ' + item.fileName);
    }

    private onFilterChanged(ev: React.FormEvent<HTMLElement>, text?: string): void {
        this.setState({
            fileLists: text ? this._allFiles.filter((item: IFileItem) => 
            hasText(item, text)) : this._allFiles
        });
    };

    private getSelectedFiles(): IFileItem[] {
        //const selectionCount = this._selection.getSelectedCount();

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

function onRenderDetailsFooter(props?: IDetailsFooterProps, defaultRender?: IRenderFunction<IDetailsFooterProps>): JSX.Element {
    return (
        <Sticky stickyPosition={StickyPositionType.Footer} isScrollSynced={true}>
            <div className={classNames.row}>
                <DetailsRow
                    columns={props!.columns}
                    item={_footerItem}
                    itemIndex={-1}
                    selection={props!.selection}
                    selectionMode={(props!.selection && props!.selection.mode) || SelectionMode.none}
                    viewport={props!.viewport}
                />
            </div>
        </Sticky>
    );
}