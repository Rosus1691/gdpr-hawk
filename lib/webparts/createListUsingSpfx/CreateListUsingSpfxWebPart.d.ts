import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ICreateListUsingSpfxWebPartProps {
    description: string;
}
export default class CreateListUsingSpfxWebPart extends BaseClientSideWebPart<ICreateListUsingSpfxWebPartProps> {
    render(): void;
    private AddEventListeners;
    private CreateListInSPOUsinPnPSPFx;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=CreateListUsingSpfxWebPart.d.ts.map