import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IKudos {
    id?: number;
    title: string;
    targetUserEmail: string;
    message: string;
    kudosType: string;
    senderName: string;
    senderEmail: string;
    kudosDate?: string;
}

export interface IKudosMapping {
    target?: string;
    message?: string;
    type?: string;
    date?: string;
}

export class KudosService {
    private readonly _context: WebPartContext;
    private readonly _listId: string;
    private readonly _mapping: IKudosMapping;

    constructor(context: WebPartContext, listId: string, mapping?: IKudosMapping) {
        this._context = context;
        this._listId = listId;
        this._mapping = mapping || {};
    }

    private get _baseUrl(): string {
        return `${this._context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this._listId}')/items`;
    }

    public async getKudosForUser(userEmail: string): Promise<IKudos[]> {
        if (!this._listId) return [];

        const targetField = this._mapping.target || "TargetUser";
        const messageField = this._mapping.message || "KudosMessage";
        const typeField = this._mapping.type || "KudosType";
        const dateField = this._mapping.date || "KudosDate";

        // Handle expansion if target field is a Person field (lookup)
        const isPersonField = targetField === "TargetUser" || targetField.includes("Id") || !targetField.includes("/");
        const expandTarget = isPersonField ? `${targetField}` : "";
        const emailProperty = isPersonField ? `${targetField}/EMail` : targetField;

        const filter = `$filter=${emailProperty} eq '${userEmail}'`;
        const select = `$select=Id,Title,${messageField},${typeField},${dateField},${targetField}/EMail,Author/Title,Author/EMail`;
        const expand = `$expand=Author${expandTarget ? `,${expandTarget}` : ""}`;

        try {
            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                `${this._baseUrl}?${select}&${expand}&${filter}`,
                SPHttpClient.configurations.v1
            );
            const data = await response.json();
            return (data.value || []).map((item: any) => ({
                id: item.Id,
                title: item.Title,
                message: item[messageField],
                kudosType: item[typeField],
                kudosDate: item[dateField] || item.Created,
                senderName: item.Author?.Title,
                senderEmail: item.Author?.EMail,
                targetUserEmail: userEmail
            }));
        } catch (error) {
            console.error("Error fetching kudos", error);
            return [];
        }
    }

    public async getKudosCounts(): Promise<Record<string, number>> {
        if (!this._listId) return {};

        const targetField = this._mapping.target || "TargetUser";
        const emailProperty = (targetField === "TargetUser" || !targetField.includes("/")) ? `${targetField}/EMail` : targetField;

        const select = `$select=Id,${targetField}/EMail`;
        const expand = `$expand=${targetField}`;

        try {
            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                `${this._baseUrl}?${select}&${expand}&$top=5000`,
                SPHttpClient.configurations.v1
            );
            const data = await response.json();
            const counts: Record<string, number> = {};
            (data.value || []).forEach((item: any) => {
                // Safely get email from nested property
                const parts = targetField.split('/');
                let current = item;
                for (const part of parts) {
                    if (current) current = current[part];
                }
                const email = current?.EMail || current; // Handle both expanded person and raw field

                if (email && typeof email === 'string') {
                    counts[email] = (counts[email] || 0) + 1;
                }
            });
            return counts;
        } catch (error) {
            console.error("Error fetching kudos counts", error);
            return {};
        }
    }

    public async submitKudos(kudos: IKudos): Promise<boolean> {
        if (!this._listId) return false;

        const messageField = this._mapping.message || "KudosMessage";
        const typeField = this._mapping.type || "KudosType";
        const targetField = this._mapping.target || "TargetUser";

        const body: any = {
            Title: kudos.title,
            [messageField]: kudos.message,
            [typeField]: kudos.kudosType,
        };

        // If it's the default TargetUser field or looks like a Person field, resolve ID
        if (targetField === "TargetUser" || !targetField.includes("/")) {
            const resolvedUser = await this._resolveUser(kudos.targetUserEmail);
            if (resolvedUser) {
                body[`${targetField}Id`] = resolvedUser;
            }
        } else {
            // If it's a text field or already mapped, try to set it directly
            body[targetField] = kudos.targetUserEmail;
        }

        try {
            const response: SPHttpClientResponse = await this._context.spHttpClient.post(
                this._baseUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=nometadata',
                        'odata-version': ''
                    },
                    body: JSON.stringify(body)
                }
            );
            return response.ok;
        } catch (error) {
            console.error("Error submitting kudos", error);
            return false;
        }
    }

    private async _resolveUser(email: string): Promise<number | null> {
        try {
            const response = await this._context.spHttpClient.post(
                `${this._context.pageContext.web.absoluteUrl}/_api/web/ensureuser`,
                SPHttpClient.configurations.v1,
                {
                    body: JSON.stringify({ logonName: email })
                }
            );
            const data = await response.json();
            return data.Id || null;
        } catch (e) {
            console.error("Error resolving user ID", e);
            return null;
        }
    }
}
