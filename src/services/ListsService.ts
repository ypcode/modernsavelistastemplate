
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { IList } from "../models/IList";
import { getAbsoluteUrl } from "../utils/utilities";
import { ContextServiceKey } from './ContextService';

export interface IListsService {
    getListAbsoluteUrl(listId: string): Promise<string>;
}

export class ListsService implements IListsService {
    private spHttpClient: SPHttpClient;
    private webUrl: string;

    constructor(serviceScope: ServiceScope) {
        this.spHttpClient = new SPHttpClient(serviceScope);
        serviceScope.whenFinished(() => {
            const context = serviceScope.consume(ContextServiceKey);
            this.webUrl = context.webUrl;
        });
    }

    public getListAbsoluteUrl(listId: string): Promise<string> {
        let url = `${this.webUrl}/_api/web/lists(guid'${encodeURIComponent(listId)}')?$expand=RootFolder`;
        return this.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response) => {
            if (response.status == 204) {
                return {};
            } else {
                return response.json();
            }
        }).then((list: IList) => getAbsoluteUrl(this.webUrl, list.RootFolder.ServerRelativeUrl));
    }
}

export const ListsServiceKey = ServiceKey.create<IListsService>(
    'YPCODE:ListsService',
    ListsService
);