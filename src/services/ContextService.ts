
import { ServiceKey } from '@microsoft/sp-core-library';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface IContextService {
    configure(context: BaseComponentContext): void;
    webUrl: string;
    isSharePointAdmin: boolean;
}

export class ContextService implements IContextService {
    private _context: BaseComponentContext;

    public configure(context: BaseComponentContext): void {
        this._context = context;
        console.log('context= ', context);
    }
    public get webUrl(): string {
        if (!this._context) {
            return null;
        }

        return this._context.pageContext.web.absoluteUrl;
    }

    public get isSharePointAdmin(): boolean {
        if (!this._context) {
            throw new Error("Context is not initialized in service");
        }

        console.log('pageContext', this._context.pageContext);
        return this._context.pageContext.legacyPageContext.isSiteAdmin;
    }
}

export const ContextServiceKey = ServiceKey.create<IContextService>(
    'YPCODE:ContextService',
    ContextService
);