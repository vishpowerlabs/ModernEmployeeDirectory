import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IEmployee } from '../models/IEmployee';

export class GraphService {
    private readonly _context: WebPartContext;
    private _client!: MSGraphClientV3;

    constructor(context: WebPartContext) {
        this._context = context;
    }

    public async getClient(): Promise<MSGraphClientV3> {
        if (!this._client) {
            this._client = await this._context.msGraphClientFactory.getClient('3');
        }
        return this._client;
    }

    public async getUsers(
        filterType?: 'none' | 'domain' | 'extension' | 'department' | 'location',
        filterValue?: string,
        filterSecondaryValue?: string,
        nextLink?: string
    ): Promise<{ users: IEmployee[], nextLink?: string }> {
        try {
            return await this._fetchUsersFromGraph(filterType, filterValue, filterSecondaryValue, nextLink);
        } catch (error) {
            console.error("Error fetching users", error);
            return { users: [], nextLink: undefined };
        }
    }

    private async _fetchUsersFromGraph(
        filterType?: 'none' | 'domain' | 'extension' | 'department' | 'location',
        filterValue?: string,
        filterSecondaryValue?: string,
        nextLink?: string
    ): Promise<{ users: IEmployee[], nextLink?: string }> {
        const client = await this.getClient();
        let response;

        if (nextLink) {
            response = await client.api(nextLink).get();
        } else {
            let request = client.api('/users')
                .select('id,displayName,jobTitle,mail,mobilePhone,businessPhones,officeLocation,department,userPrincipalName,companyName,city,state,country,onPremisesExtensionAttributes')
                .top(50); // Pagination page size

            // Apply Server-Side Filters
            if (filterType && filterType !== 'none' && filterValue) {
                if (filterType === 'department') {
                    request = request.filter(`startsWith(department, '${filterValue}')`);
                } else if (filterType === 'location') {
                    request = request.filter(`startsWith(officeLocation, '${filterValue}')`);
                } else if (filterType === 'extension' && filterSecondaryValue) {
                    request = request.filter(`onPremisesExtensionAttributes/${filterValue} eq '${filterSecondaryValue}'`);
                }
                // Note: 'domain' filtering is complex server-side without advanced queries. 
                // For now, if domain filter is valid, we might need to fetch more or use client-side filtering on the page.
                // However, sticking to the requested server-side scope:
            }

            response = await request.get();
        }

        let users: IEmployee[] = response.value;
        const responseNextLink = response['@odata.nextLink'];

        // Client-side fallback for 'domain' filter (since we can't easily exact match domain in standard filter)
        if (filterType === 'domain' && filterValue) {
            const searchVal = filterValue.toLowerCase();
            users = users.filter(u =>
                u.mail?.toLowerCase().endsWith(`@${searchVal}`) ||
                u.userPrincipalName?.toLowerCase().endsWith(`@${searchVal}`)
            );
            // Note: Pagination breaks here if we filter client-side after fetching a page. 
            // In a real scenario, we'd use $search or get all. 
            // For this implementation, let's accept that 'domain' filter checks the current page.
        }

        return { users: users, nextLink: responseNextLink };
    }

    // Cache methods removed for pagination simplicity in this iteration


    public async getUserPhoto(userId: string): Promise<string | undefined> {
        const client = await this.getClient();
        try {
            const response = await client.api(`/users/${userId}/photo/$value`)
                .responseType('blob' as any)
                .get();

            const url = (globalThis.URL as any)?.createObjectURL ? globalThis.URL : (globalThis as any).webkitURL;
            return url.createObjectURL(response);
        } catch (error) {
            console.warn(`Could not fetch photo for user ${userId}. Falling back to initials.`, error);
            return undefined;
        }
    }

    public async getCurrentUser(): Promise<IEmployee | undefined> {
        const client = await this.getClient();
        try {
            const user = await client.api('/me').select('id,displayName,jobTitle,mail,userPrincipalName').get();
            return user;
        } catch (error) {
            console.error("Error fetching current user", error);
            return undefined;
        }
    }
    public async getManager(userId: string): Promise<IEmployee | null> {
        try {
            const client = await this.getClient();
            const manager = await client.api(`/users/${userId}/manager`).select('id,displayName,jobTitle,mail,department').get();
            return manager;
        } catch (error) {
            console.warn(`No manager found for user ${userId}`, error);
            return null;
        }
    }

    public async getDirectReports(userId: string): Promise<IEmployee[]> {
        try {
            const client = await this.getClient();
            const response = await client.api(`/users/${userId}/directReports`).select('id,displayName,jobTitle,mail,department,officeLocation').get();
            return response.value;
        } catch (error) {
            console.warn(`Error fetching direct reports for ${userId}`, error);
            return [];
        }
    }

    public async getPresence(userIds: string[]): Promise<any[]> {
        if (!userIds || userIds.length === 0) return [];

        try {
            const client = await this.getClient();
            const response = await client.api('/communications/getPresencesByUserId')
                .post({
                    ids: userIds
                });
            return response.value;
        } catch (error) {
            console.warn("Error fetching presence", error);
            return [];
        }
    }

    public async getPeopleIWorkWith(userId: string): Promise<IEmployee[]> {
        try {
            const client = await this.getClient();
            // Using /users/{id}/people which returns people relevant to the user (Insights)
            const response = await client.api(`/users/${userId}/people`)
                .select('id,displayName,jobTitle,userPrincipalName,department')
                .top(10)
                .get();

            // Map to compatible IEmployee interface
            return response.value.map((p: any) => ({
                id: p.id,
                displayName: p.displayName,
                jobTitle: p.jobTitle,
                mail: p.userPrincipalName, // Most people have mail same as UPN or we can use UPN
                department: p.department
            }));
        } catch (error) {
            console.warn(`Error fetching people for ${userId}`, error);
            return [];
        }
    }

    public async getUserDetails(userId: string): Promise<Partial<IEmployee>> {
        try {
            const client = await this.getClient();
            const response = await client.api(`/users/${userId}`)
                .version('beta')
                .select('aboutMe,birthday,hireDate,skills,interests,schools,pastProjects,mobilePhone,businessPhones,officeLocation')
                .get();
            return response;
        } catch (error) {
            console.warn(`Error fetching detailed profile for ${userId}`, error);
            return {};
        }
    }

    public async updateMyProfile(updates: Partial<IEmployee>): Promise<void> {
        try {
            const client = await this.getClient();
            // Using /beta/me to support aboutMe, skills, interests and pastProjects
            await client.api('/me')
                .version('beta')
                .patch(updates);
        } catch (error) {
            console.error("Error updating profile", error);
            throw error;
        }
    }

    public async getAvailableProperties(): Promise<string[]> {
        try {
            const client = await this.getClient();
            // Fetch the current user as a representative sample
            const response = await client.api('/me').get();
            console.log("DEBUG: Graph /me response for discovery:", response);

            // Extract all keys that have values (including arrays)
            const keys = Object.keys(response).filter(key =>
                !key.startsWith('@') &&
                (typeof response[key] !== 'object' || Array.isArray(response[key])) &&
                response[key] !== null
            );

            // Add some standard ones that might be missing in /me but common
            const defaults = ['displayName', 'jobTitle', 'mail', 'mobilePhone', 'businessPhones', 'officeLocation', 'department', 'companyName', 'city', 'state', 'country'];
            const allKeys = Array.from(new Set([...defaults, ...keys])).sort((a, b) => a.localeCompare(b));

            console.log("DEBUG: Final Discovered Properties:", allKeys);
            return allKeys;
        } catch (error) {
            console.error("Error discovering properties", error);
            return ['displayName', 'jobTitle', 'mail', 'mobilePhone', 'businessPhones', 'officeLocation', 'department', 'companyName', 'city', 'state', 'country'];
        }
    }
    public async getManagers(userId: string): Promise<IEmployee[]> {
        const managers: IEmployee[] = [];
        let currentUserId = userId;
        let depth = 0;
        const maxDepth = 5;

        try {
            const client = await this.getClient();

            while (depth < maxDepth) {
                try {
                    const manager = await client.api(`/users/${currentUserId}/manager`)
                        .select('id,displayName,jobTitle,mail')
                        .get();

                    if (manager) {
                        managers.unshift(manager); // Add to beginning (highest boss first)
                        currentUserId = manager.id;
                        depth++;
                    } else {
                        break;
                    }
                } catch (e) {
                    // 404 means no manager or permissions issue - gracefully stop climbing
                    console.warn(`Manager fetch stopped for user ${currentUserId}. Reason: ${e instanceof Error ? e.message : 'Unknown error'}`);
                    break;
                }
            }
        } catch (error) {
            console.error("Error fetching managers", error);
        }

        return managers;
    }
}
