import * as React from 'react';

import { DefaultButton, Stack, ProgressIndicator, Icon } from 'office-ui-fabric-react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import {
    DetailsList,
    DetailsListLayoutMode,
    IDetailsHeaderProps,
    Selection,
    IColumn,
    ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { IFileItem, ItemList } from './ItemList';
import { classNames } from './ComponentStyles';

export interface IAttachmentProps {
    regardingObjectId: string;
    regardingEntityName: string;
    files: IFileItem[];
    onAttach: (selectedFiles: IFileItem[]) => Promise<void>;
}

export interface IAttachmentState {
    files: IFileItem[];
    columns: IColumn[];
    hiddenModal: boolean;
    isInProgress: boolean;
}

export class AttachmentManagerApp extends React.Component<IAttachmentProps, IAttachmentState> {
    private selection: Selection;
    private allFiles: ItemList;

    constructor(props: IAttachmentProps) {
        super(props)

        initializeIcons();

        this.allFiles = new ItemList();

        this.selection = new Selection();

        this.allFiles.setItems(this.props.files);
        
        this.state = {
            files: this.allFiles.getItems(),
            hiddenModal: true,
            isInProgress: false,
            columns: this.allFiles.getColumns()
        };

        this.attachFilesClicked = this.attachFilesClicked.bind(this);
        this.onFilterChanged = this.onFilterChanged.bind(this);
        this.onAttachClicked = this.onAttachClicked.bind(this);
        this.hideDialog = this.hideDialog.bind(this);
    }

    public render(): JSX.Element {
        const { hiddenModal: hiddenDialog, files, columns } = this.state;
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
                    minWidth='900px'
                >
                    <div className={classNames.wrapper}>
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
                            <MarqueeSelection selection={this.selection}>
                                <DetailsList
                                    items={files}
                                    columns={columns}
                                    setKey="set"
                                    layoutMode={DetailsListLayoutMode.fixedColumns}
                                    constrainMode={ConstrainMode.unconstrained}
                                    onRenderItemColumn={renderItemColumn}
                                    onRenderDetailsHeader={onRenderDetailsHeader}
                                    selection={this.selection}
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
        console.log('Item invoked: ' + item.fileName);
    }

    private onFilterChanged(ev?: React.ChangeEvent<HTMLInputElement>, text?: string): void {
        this.setState({
            files: text ? this.state.files.filter((item: IFileItem) => 
            hasText(item, text)) : this.state.files
        });
    };

    private getSelectedFiles(): IFileItem[] {
        let selectedFiles: IFileItem[] = [];

        for(let i = 0; i < this.selection.getSelectedCount(); i++) {
            selectedFiles.push((this.selection.getSelection()[i] as IFileItem));
        }

        return selectedFiles;
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

function renderItemColumn(item: IFileItem, index?: number, column?: IColumn) {
    if (column) {
        const fieldContent = item[column.fieldName as keyof IFileItem] as string;

        switch (column.key) {
            case 'iconclassname':
                return <Icon iconName={fieldContent} className={classNames.fileIcon}></Icon>;
            case 'lastModifiedOn':
                const dateField = item[column.fieldName as keyof IFileItem] as Date;
                return <div>{dateField.toLocaleDateString('en-nz')} {dateField.toLocaleTimeString('en-nz')}</div>;
            case 'fileType':
            case 'fileName':
            case 'lastModifiedBy':
                return <div>{fieldContent}</div>;
            default:
                break;
        }
    }
}