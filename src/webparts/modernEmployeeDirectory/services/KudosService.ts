import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IKudos, IKudosListItem } from '../models/IKudos';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';

export interface IKudosServiceConfig {
    listId: string;
    recipientColumn: string;
    authorColumn: string;
    messageColumn: string;
    badgeTypeColumn: string;
}

export class KudosService {
    private readonly sp: SPFI;
    private readonly config: IKudosServiceConfig;
    private resolvedConfig: IKudosServiceConfig | null = null;
    private readonly resolutionPromise: Promise<void>;

    constructor(context: WebPartContext, config: IKudosServiceConfig) {
        this.sp = spfi().using(SPFx(context));
        this.config = config;
        // Start resolution and store the promise (initialized via sync call)
        this.resolutionPromise = this._startResolution();
    }

    private _startResolution(): Promise<void> {
        return this.resolveColumnNames();
    }

    private async resolveColumnNames(): Promise<void> {
        try {
            console.log('[KudosService] Resolving column names from GUIDs...');
            console.log('[KudosService] Original config:', this.config);

            const resolvedConfig: IKudosServiceConfig = {
                listId: this.config.listId,
                recipientColumn: this.config.recipientColumn,
                authorColumn: this.config.authorColumn,
                messageColumn: this.config.messageColumn,
                badgeTypeColumn: this.config.badgeTypeColumn
            };

            // Check if columns are GUIDs (contain hyphens) and resolve them
            const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

            if (guidPattern.test(this.config.recipientColumn)) {
                const field = await this.sp.web.lists.getById(this.config.listId).fields.getById(this.config.recipientColumn)();
                resolvedConfig.recipientColumn = field.InternalName;
                console.log('[KudosService] Resolved recipientColumn:', field.InternalName);
            }

            if (guidPattern.test(this.config.authorColumn)) {
                const field = await this.sp.web.lists.getById(this.config.listId).fields.getById(this.config.authorColumn)();
                resolvedConfig.authorColumn = field.InternalName;
                console.log('[KudosService] Resolved authorColumn:', field.InternalName);
            }

            if (guidPattern.test(this.config.messageColumn)) {
                const field = await this.sp.web.lists.getById(this.config.listId).fields.getById(this.config.messageColumn)();
                resolvedConfig.messageColumn = field.InternalName;
                console.log('[KudosService] Resolved messageColumn:', field.InternalName);
            }

            if (guidPattern.test(this.config.badgeTypeColumn)) {
                const field = await this.sp.web.lists.getById(this.config.listId).fields.getById(this.config.badgeTypeColumn)();
                resolvedConfig.badgeTypeColumn = field.InternalName;
                console.log('[KudosService] Resolved badgeTypeColumn:', field.InternalName);
            }

            this.resolvedConfig = resolvedConfig;
            console.log('[KudosService] ✅ Resolution complete. Resolved config:', this.resolvedConfig);
        } catch (error) {
            console.error('[KudosService] Error resolving column names:', error);
            // Fall back to original config
            this.resolvedConfig = this.config;
        }
    }

    private async getConfig(): Promise<IKudosServiceConfig> {
        // Wait for resolution to complete before returning config
        await this.resolutionPromise;
        return this.resolvedConfig || this.config;
    }

    public async getBadgeTypeChoices(): Promise<string[]> {
        try {
            const config = await this.getConfig();

            if (!config.listId || !config.badgeTypeColumn) {
                console.warn('[KudosService] List ID or badge type column not configured');
                return [];
            }

            // Fetch the field definition to get choices
            const field = await this.sp.web.lists
                .getById(config.listId)
                .fields
                .getByInternalNameOrTitle(config.badgeTypeColumn)();

            // Check if it's a choice field
            if (field.FieldTypeKind === 6) { // 6 = Choice field
                const choiceField = field as any;
                const choices = choiceField.Choices || [];
                console.log('[KudosService] Badge type choices:', choices);
                return choices;
            }

            console.warn('[KudosService] Badge type column is not a choice field');
            return [];
        } catch (error) {
            console.error('[KudosService] Error fetching badge type choices:', error);
            return [];
        }
    }

    public async getKudosForUser(userId: string): Promise<IKudos[]> {
        try {
            console.log('[KudosService] getKudosForUser called with userId:', userId);
            const config = await this.getConfig();
            console.log('[KudosService] Using config:', config);

            if (!config.listId) {
                console.warn('[KudosService] No list ID configured');
                return [];
            }

            // First, ensure the user and get their ID
            const user = await this.sp.web.ensureUser(userId);
            console.log('[KudosService] Resolved user:', { Id: user.Id, LoginName: user.LoginName, Email: user.Email });

            const filterQuery = `${config.recipientColumn}Id eq ${user.Id}`;
            console.log('[KudosService] Filter query:', filterQuery);

            const items = await this.sp.web.lists
                .getById(config.listId)
                .items
                .filter(filterQuery)
                .expand('Author')
                .select('Id', 'Created', config.messageColumn, config.badgeTypeColumn, 'Author/Title', 'Author/EMail')
                .orderBy('Created', false)();

            console.log('[KudosService] Fetched items count:', items.length);
            console.log('[KudosService] Fetched items:', items);

            return items.map((item: any) => ({
                id: item.Id.toString(),
                recipientId: userId,
                recipientName: '',
                authorId: item.Author?.EMail || 'unknown',
                authorName: item.Author?.Title || 'Unknown',
                message: item[config.messageColumn] || '',
                badgeType: item[config.badgeTypeColumn] || 'star',
                createdDate: new Date(item.Created)
            }));
        } catch (error) {
            console.error('[KudosService] Error fetching kudos:', error);
            return [];
        }
    }

    public async getKudosCount(userId: string): Promise<number> {
        try {
            console.log('[KudosService] Fetching kudos count for user:', userId);
            const config = await this.getConfig();

            const user = await this.sp.web.ensureUser(userId);
            console.log('[KudosService] Resolved user ID:', user.Id);

            const items = await this.sp.web.lists
                .getById(config.listId)
                .items
                .filter(`${config.recipientColumn}Id eq ${user.Id}`)
                .select('Id')();

            console.log('[KudosService] Kudos count for', userId, ':', items.length);
            return items.length;
        } catch (error) {
            console.error('[KudosService] Error fetching kudos count for', userId, ':', error);
            return 0;
        }
    }

    /**
     * Finds users who have received at least minCount kudos.
     * Returns an array of user principal names (emails).
     */
    public async getHallOfFameCandidates(minCount: number): Promise<string[]> {
        if (minCount <= 0) return [];

        try {
            console.log('[KudosService] Finding candidates with min kudos:', minCount);
            const config = await this.getConfig();

            if (!config.listId) return [];

            // Fetch all kudos recipients. Since SP doesn't support GROUP BY easily via REST,
            // we'll fetch ID and RecipientEmail and aggregate in JS.
            const items = await this.sp.web.lists
                .getById(config.listId)
                .items
                .select(config.recipientColumn + '/EMail')
                .expand(config.recipientColumn)();

            const counts: { [email: string]: number } = {};
            items.forEach((item: any) => {
                const email = item[config.recipientColumn]?.EMail;
                if (email) {
                    counts[email] = (counts[email] || 0) + 1;
                }
            });

            const candidates = Object.keys(counts).filter(email => counts[email] >= minCount);
            console.log('[KudosService] Found Hall of Fame candidates:', candidates);
            return candidates;

        } catch (error) {
            console.error('[KudosService] Error finding Hall of Fame candidates:', error);
            return [];
        }
    }

    public async giveKudos(recipientId: string, recipientName: string, message: string, badgeType: string): Promise<boolean> {
        try {
            console.log('[KudosService] giveKudos called:', { recipientId, recipientName, message, badgeType });
            const config = await this.getConfig();
            console.log('[KudosService] Using config for giveKudos:', config);

            // Ensure the recipient user exists and get their ID
            const recipientUser = await this.sp.web.ensureUser(recipientId);
            console.log('[KudosService] Resolved recipient user ID:', recipientUser.Id);

            const itemData: any = {
                [config.messageColumn]: message,
                [config.badgeTypeColumn]: badgeType
            };

            // Set the recipient using the lookup field format
            itemData[`${config.recipientColumn}Id`] = recipientUser.Id;

            console.log('[KudosService] Creating item with data:', itemData);

            await this.sp.web.lists
                .getById(config.listId)
                .items.add(itemData);

            console.log('[KudosService] Kudos saved successfully');
            return true;
        } catch (error) {
            console.error('[KudosService] Error giving kudos:', error);
            return false;
        }
    }

    private mapToKudos(item: IKudosListItem): IKudos {
        return {
            id: item.Id.toString(),
            recipientId: item[this.config.recipientColumn] || '',
            recipientName: item[this.config.recipientColumn] || '',
            authorId: item[this.config.authorColumn] || '',
            authorName: item[this.config.authorColumn] || '',
            message: item[this.config.messageColumn] || '',
            badgeType: item[this.config.badgeTypeColumn] || 'star',
            createdDate: new Date(item.Created)
        };
    }
}
