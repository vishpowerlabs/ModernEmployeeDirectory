import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/fields';

export interface IAuditServiceConfig {
    listId: string;
    activityColumn: string;
    actorColumn: string;
    targetColumn: string;
    detailsColumn: string;
    onLog?: (message: string, type: 'info' | 'success' | 'error' | 'warning') => void;
}

export class AuditService {
    private readonly sp: SPFI;
    private readonly config: IAuditServiceConfig;
    private _columnCache: { [guid: string]: string } = {};

    constructor(context: WebPartContext, config: IAuditServiceConfig) {
        this.sp = spfi().using(SPFx(context));
        this.config = config;
        this._safeLog('AuditService v2 Init.', 'info');
    }

    private _isGuid(str: string): boolean {
        return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(str);
    }

    private async _resolveColumn(columnName: string): Promise<string> {
        if (!columnName || !this._isGuid(columnName)) return columnName;
        if (this._columnCache[columnName]) return this._columnCache[columnName];

        try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const field: any = await this.sp.web.lists.getById(this.config.listId).fields.getById(columnName).select('InternalName')();
            if (field && field.InternalName) {
                this._columnCache[columnName] = field.InternalName;
                return field.InternalName;
            }
            return columnName;
        } catch (error) {
            console.error('[AuditService] Error resolving column:', columnName, error);
            return columnName;
        }
    }

    private _safeLog(message: string, type: 'info' | 'success' | 'error' | 'warning'): void {
        const prefix = `[AUDITLOG]`;
        if (type === 'error') console.error(`${prefix} ${message}`);

        if (this.config.onLog) {
            this.config.onLog(message, type);
        }
    }

    public async logActivity(activity: string, target: string, details: string, actorId?: string): Promise<boolean> {
        try {
            if (!this.config.listId) return false;

            // Step 1: Resolve columns
            const activityCol = await this._resolveColumn(this.config.activityColumn);
            const targetCol = await this._resolveColumn(this.config.targetColumn);
            const detailsCol = await this._resolveColumn(this.config.detailsColumn);
            const actorCol = await this._resolveColumn(this.config.actorColumn);

            const itemData: Record<string, string | number> = {
                [activityCol]: activity,
                [targetCol]: target,
                [detailsCol]: details
            };

            // Step 2: Resolve Actor ID (optional)
            if (actorId && actorCol) {
                try {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const userResult: any = await this.sp.web.ensureUser(actorId);
                    // Use a safe multi-step ID extraction
                    let finalId: number | undefined;
                    if (userResult) {
                        if (userResult.data) finalId = userResult.data.Id;
                        if (!finalId) finalId = userResult.Id;
                    }

                    if (finalId) {
                        itemData[`${actorCol}Id`] = finalId;
                    } else {
                        itemData[actorCol] = actorId;
                    }
                } catch (e) {
                    console.log('[AuditService] Could not ensure user, using actorId as text:', actorId, e);
                    itemData[actorCol] = actorId;
                }
            }

            // Step 3: Add Item
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const result: any = await this.sp.web.lists.getById(this.config.listId).items.add(itemData);

            // Step 4: Confirm Success
            let createdId: number | undefined;
            if (result) {
                if (result.data) createdId = result.data.Id;
                if (!createdId) createdId = result.Id;
            }

            if (createdId) {
                this._safeLog(`SUCCESS: ${activity} (ID: ${createdId})`, 'success');
            } else {
                this._safeLog(`SENT: ${activity}`, 'info');
            }

            return true;
        } catch (err) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const msg = (err as any)?.message || 'Unknown';
            this._safeLog(`[SERVICE-ERROR] Activity failed: ${msg}`, 'error');
            return false;
        }
    }
}
