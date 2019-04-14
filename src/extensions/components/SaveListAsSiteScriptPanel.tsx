import * as React from "react";
import { Panel, PanelType, PrimaryButton, TextField, DialogFooter, Toggle, assign, SpinnerType, Spinner, Pivot, PivotLinkFormat, PivotItem, Dropdown, IDropdownOption, DefaultButton } from "office-ui-fabric-react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { SiteDesignsServiceKey, ISiteDesignsService } from "../../services/SiteDesignsService";
import { IListsService, ListsServiceKey } from "../../services/ListsService";
import { ISiteScript, ISiteScriptContent, ISiteScriptAction } from "../../models/ISiteScript";
import { download } from "../../utils/utilities";
import styles from "./SaveListAsSiteScriptPanel.module.scss";
import { IContextService, ContextServiceKey } from "../../services/ContextService";
import { WebTemplate, ISiteDesign } from "../../models/ISiteDesign";


export enum OperationStatus {
    None,
    Success,
    Error
}

export enum WizardStep {
    TemplateSettings,
    AssociateToSiteDesign
}

export interface ISaveListAsSiteScriptPanelProps {
    isOpen: boolean;
    listId: string;
    listTitle: string;
    onClose?: () => void;
    serviceScope: ServiceScope;
}

export interface ISaveListAsSiteScriptPanelState {
    isLoading: boolean;
    isSaving: boolean;
    templateTitle: string;
    templateDescription: string;
    includeDescription: boolean;
    includeContentTypes: boolean;
    includeViews: boolean;
    includeNavLink: boolean;
    useExistingSiteDesign: boolean;
    newSiteDesignTitle: string;
    newSiteDesignDescription: string;
    newSiteDesignWebTemplate: WebTemplate;
    existingSiteDesigns: ISiteDesign[];
    selectedExistingSiteDesignId: string;
    userMessage: string;
    operationStatus: OperationStatus;
    wizardStep: WizardStep;
    savedSiteScriptId: string;
}

export class SaveListAsSiteScriptPanel extends React.Component<ISaveListAsSiteScriptPanelProps, ISaveListAsSiteScriptPanelState> {

    private siteDesignsService: ISiteDesignsService;
    private listsService: IListsService;
    private contextService: IContextService;

    constructor(props: ISaveListAsSiteScriptPanelProps) {
        super(props);

        if (props.serviceScope == null) {
            throw new Error("The service scope instance has not been passed");
        }

        this.state = {
            isLoading: true,
            isSaving: false,
            templateTitle: '',
            templateDescription: '',
            includeDescription: true,
            includeContentTypes: true,
            includeViews: true,
            includeNavLink: true,
            useExistingSiteDesign: true,
            newSiteDesignTitle: '',
            newSiteDesignDescription: '',
            newSiteDesignWebTemplate: WebTemplate.None,
            userMessage: '',
            operationStatus: OperationStatus.None,
            wizardStep: WizardStep.TemplateSettings,
            existingSiteDesigns: [],
            selectedExistingSiteDesignId: null,
            savedSiteScriptId: null,
        };
    }

    public componentWillMount() {
        // Ensure the service scope is initialized
        this.props.serviceScope.whenFinished(() => {
            this.siteDesignsService = this.props.serviceScope.consume(SiteDesignsServiceKey);
            this.listsService = this.props.serviceScope.consume(ListsServiceKey);
            this.contextService = this.props.serviceScope.consume(ContextServiceKey);

            this.siteDesignsService.getSiteDesigns().then(siteDesigns => {
                this.setState({ isLoading: false, existingSiteDesigns: siteDesigns });
            }).catch(err => {
                this.setState({ isLoading: false });
            });
        });
    }

    private _onClose(): void {
        if (this.props.onClose) {
            this.props.onClose();
        }
    }

    private _renderTemplateSettingsForm() {
        let { listTitle } = this.props;
        let { templateDescription, templateTitle, includeContentTypes,
            includeDescription, includeNavLink, includeViews,
            isSaving, operationStatus, userMessage } = this.state;

        return (<div className={styles.saveListAsSiteScript} >
            <h1>Save list {listTitle} as template</h1>
            <h2>Template</h2>
            <div className={styles.row}>
                <div className={styles.column}>
                    <TextField value={templateTitle} label="Name" onChanged={v => this.setState({ templateTitle: v })} />
                </div>
            </div>
            <div className={styles.row}>
                <div className={styles.column}>
                    <TextField value={templateDescription} label="Description" multiline={true} rows={6} onChanged={v => this.setState({ templateDescription: v })} />
                </div>
            </div>
            <h2>Options</h2>
            <div className={styles.row}>
                <div className={`${styles.column}`}>
                    <Toggle checked={includeViews} label="Include views" onChanged={v => this.setState({ includeViews: v })} />
                </div>
                <div className={`${styles.column}`}>
                    <Toggle checked={includeContentTypes} label="Include Content Types" onChanged={v => this.setState({ includeContentTypes: v })} />
                </div>
            </div>
            <div className={styles.row}>
                <div className={`${styles.column}`}>
                    <Toggle checked={includeNavLink} label="Add to navigation" onChanged={v => this.setState({ includeNavLink: v })} />
                </div>
                <div className={`${styles.column}`}>
                    <Toggle checked={includeDescription} label="Include description" onChanged={v => this.setState({ includeDescription: v })} />
                </div>
            </div>
            <div>
                {isSaving && <div><br />
                    <Spinner type={SpinnerType.large} label="Saving..." />
                    <br />
                </div>}
                {userMessage && <div className={operationStatus == OperationStatus.Success
                    ? styles.success
                    : (operationStatus == OperationStatus.Error
                        ? styles.error
                        : '')}>
                    {userMessage}
                </div>}
            </div>
            <DialogFooter>
                <PrimaryButton iconProps={{ iconName: 'Download' }} text="Download" onClick={() => this._saveAsFile()} />
                {this.contextService.isSharePointAdmin && <PrimaryButton iconProps={{ iconName: 'Save' }} text="Save" onClick={() => this._saveAsTenantSiteScript()} />}
            </DialogFooter>

        </div>
        );
    }

    private _renderNewSiteDesignForm() {
        let { newSiteDesignTitle, newSiteDesignDescription, newSiteDesignWebTemplate } = this.state;
        return (<div>
            <div className={styles.row}>
                <div className={styles.column}>
                    <TextField value={newSiteDesignTitle} label="Site Design name"
                        onChanged={v => this.setState({ newSiteDesignTitle: v })} />
                </div>
            </div>
            <div className={styles.row}>
                <div className={styles.column}>
                    <TextField value={newSiteDesignDescription} label="Site Design description" multiline={true} rows={6}
                        onChanged={v => this.setState({ newSiteDesignDescription: v })} />
                </div>
            </div>
            <div className={styles.row}>
                <div className={styles.column}>
                    <Dropdown selectedKey={newSiteDesignWebTemplate}
                        label="Web Template"
                        onChanged={(o) => this.setState({ newSiteDesignWebTemplate: o.key as WebTemplate })}
                        options={[
                            {
                                key: WebTemplate.TeamSite,
                                text: "Team Site"
                            },
                            {
                                key: WebTemplate.CommunicationSite,
                                text: "Communication Site"
                            }]} />
                </div>
            </div>
        </div>);
    }

    private _getExistingSiteDesignOptions(): IDropdownOption[] {
        return this.state.existingSiteDesigns.map(sd => ({ key: sd.Id, text: sd.Title }));
    }

    private _renderAssociateToSiteDesign() {
        let { useExistingSiteDesign, isSaving, userMessage, operationStatus } = this.state;
        return (<div className={styles.saveListAsSiteScript}>
            <p>The Site Script has been saved to your tenant.
                It cannot be used until it is associated to a Site Design.
                <br />
                <br />
                Do you want to associate it to a site design ?</p>
            <div className={styles.row}>
                <div className={styles.column}>
                    <Pivot linkFormat={PivotLinkFormat.tabs}
                        selectedKey={useExistingSiteDesign ? 'existing' : 'new'}
                        onLinkClick={(item) => this.setState({ useExistingSiteDesign: item.props.itemKey != 'new' })}>
                        <PivotItem headerText="Existing Site Design" itemKey='existing'>
                            <Dropdown options={this._getExistingSiteDesignOptions()} onChanged={(o => this.setState({ selectedExistingSiteDesignId: o.key.toString() }))} label="Choose the Site Design" />
                        </PivotItem>
                        <PivotItem headerText="New Site Design" itemKey='new'>
                            {this._renderNewSiteDesignForm()}
                        </PivotItem>
                    </Pivot>
                </div>
            </div>
            <div>
                {isSaving && <div><br />
                    <Spinner type={SpinnerType.large} label="Saving..." />
                    <br />
                </div>}
                {userMessage && <div className={operationStatus == OperationStatus.Success
                    ? styles.success
                    : (operationStatus == OperationStatus.Error
                        ? styles.error
                        : '')}>
                    {userMessage}
                </div>}
            </div>
            <DialogFooter>
                <PrimaryButton iconProps={{ iconName: 'Save' }} text="Associate" onClick={() => this._associateToSiteDesign()} />
                <DefaultButton text="Cancel" onClick={() => this._onClose()} />
            </DialogFooter>
        </div>);
    }

    public render(): React.ReactElement<ISaveListAsSiteScriptPanelProps> {
        let { isOpen } = this.props;
        let { isLoading, wizardStep } = this.state;

        if (isLoading) {
            return <Panel isOpen={isOpen} type={PanelType.smallFixedFar} onDismiss={() => this._onClose()}>
                <Spinner type={SpinnerType.large} label="Loading..." ></Spinner>
            </Panel>;
        }

        return <Panel isOpen={isOpen} type={PanelType.smallFixedFar} onDismiss={() => this._onClose()}>
            {wizardStep == WizardStep.TemplateSettings && this._renderTemplateSettingsForm()}
            {wizardStep == WizardStep.AssociateToSiteDesign && this._renderAssociateToSiteDesign()}
        </Panel >;
    }

    private _setWizardStep(wizardStep: WizardStep): void {
        this.setState({
            userMessage: '',
            wizardStep
        });
    }

    private _getSiteScriptFromList(): Promise<ISiteScript> {
        return this.listsService.getListAbsoluteUrl(this.props.listId)
            .then(listUrl => this.siteDesignsService.getSiteScriptFromList(listUrl))
            .then((siteScript: ISiteScriptContent) => {
                console.log('original site script content: ', siteScript);
                const processed = this._processSiteScriptContent(siteScript);
                console.log('processed site script content: ', processed);
                let script: ISiteScript = {
                    Content: processed,
                    Description: this.state.templateDescription,
                    Title: this.state.templateTitle,
                    Version: 1
                };

                return script;
            });
    }

    private _saveAsFile(): void {
        this.setState({ isSaving: true });
        this._getSiteScriptFromList()
            .then(siteScript => {
                let fileName = `${this.state.templateTitle || "site-script"}.json`;
                // TODO Save to file and stream file download
                download(fileName, JSON.stringify(siteScript, null, 4));
                this.setState({ isSaving: false });
                this._askDelayedClose();
            }).catch(err => {
                this.setState({
                    isSaving: false,
                    userMessage: `An error occured: ${err}`,
                    operationStatus: OperationStatus.Error
                });
            });
    }

    private _saveAsTenantSiteScript(): void {
        this.setState({ isSaving: true });
        this._getSiteScriptFromList()
            .then(siteScript => {
                // Save to SharePoint
                this.siteDesignsService.saveSiteScript(siteScript)
                    .then((result) => {
                        this.setState({
                            isSaving: false,
                            userMessage: `The template has been saved`,
                            operationStatus: OperationStatus.Success,
                            newSiteDesignTitle: `${this.props.listTitle} template`,
                            savedSiteScriptId: result.Id
                        });
                        this._setWizardStep(WizardStep.AssociateToSiteDesign);
                    }).catch(err => {
                        this.setState({
                            isSaving: false,
                            userMessage: `An error occured: ${err}`,
                            operationStatus: OperationStatus.Error
                        });
                    });
            });
    }

    private _associateToSiteDesign(): void {
        let { useExistingSiteDesign,
            selectedExistingSiteDesignId,
            newSiteDesignTitle: newSiteDesignName,
            newSiteDesignDescription,
            newSiteDesignWebTemplate,
            savedSiteScriptId } = this.state;

        this.setState({ isSaving: true });
        let siteDesignPromise: Promise<ISiteDesign> = useExistingSiteDesign
            // Get the selected existing site design
            ? this.siteDesignsService.getSiteDesign(selectedExistingSiteDesignId)
            // Create a new Site Design
            : this.siteDesignsService.saveSiteDesign({
                Title: newSiteDesignName,
                Description: newSiteDesignDescription,
                WebTemplate: newSiteDesignWebTemplate.toString(),
                PreviewImageAltText: null,
                PreviewImageUrl: null
            } as ISiteDesign).then(created => this.siteDesignsService.getSiteDesign(created.Id));

        siteDesignPromise.then(siteDesign => {
            // Add the site script Id
            if (!siteDesign.SiteScriptIds) {
                siteDesign.SiteScriptIds = [];
            }
            siteDesign.SiteScriptIds.push(savedSiteScriptId);

            // Update the site design
            this.siteDesignsService.saveSiteDesign(siteDesign).then(() => {
                this.setState({
                    isSaving: false,
                    operationStatus: OperationStatus.Success,
                    userMessage: 'The Site Design has been saved.'
                });
                this._askDelayedClose();
            }).catch(err => {
                this.setState({
                    isSaving: false,
                    operationStatus: OperationStatus.Error,
                    userMessage: `An error occured. ${err}`
                });
            });
        });
    }

    private _askDelayedClose(): void {
        setTimeout(() => {
            this._onClose();
        }, 4000);
    }

    private _processSiteScriptContent(siteScript: ISiteScriptContent): ISiteScriptContent {
        let processed = assign({}, siteScript) as ISiteScriptContent;

        if (siteScript.actions) {
            let actions = siteScript.actions as ISiteScriptAction[];
            processed.actions = this._filterActions(actions);
        }

        return processed;
    }

    private _filterActions(actions: ISiteScriptAction[]): ISiteScriptAction[] {
        let { includeContentTypes, includeDescription, includeNavLink, includeViews } = this.state;
        if (actions) {

            return actions
                .map(a => assign({}, a) as ISiteScriptAction)
                .filter(a => {
                    switch (a.verb.toLowerCase()) {
                        case "addnavlink":
                            return includeNavLink;
                        case "createsplist":
                            if (a.subactions) {
                                a.subactions = this._filterActions(a.subactions);
                            }
                            return true;
                        case "addcontenttype":
                        case "removecontenttype":
                        case "createsitecolumnxml":
                        case "createcontenttype":
                        case "addsitecolumn":
                            return includeContentTypes;
                        case "setdescription":
                            return includeDescription;
                        case "addspview":
                        case "removespview":
                            return includeViews;
                        default:
                            return true;
                    }
                });
        }

        return [];
    }
} 