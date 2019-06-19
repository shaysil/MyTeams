import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
import { ITenant } from '../../shared/interfaces';
export interface IMyTeamsWebPartProps {
    tenantInfo: ITenant;
    openInClientApp: boolean;
}
export default class MyTeamsWebPart extends BaseClientSideWebPart<IMyTeamsWebPartProps> {
    private _graphClient;
    onInit(): Promise<void>;
    render(): Promise<void>;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    private _getTenantInfo;
}
