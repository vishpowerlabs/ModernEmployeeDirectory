import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphUser {
    id: string;
    displayName: string;
    givenName?: string;
    surname?: string;
    jobTitle?: string;
    department?: string;
    officeLocation?: string;
    city?: string;
    state?: string;
    country?: string;
    mail?: string;
    userPrincipalName?: string;
    mobilePhone?: string;
    businessPhones?: string[];
    // Extended profile fields
    aboutMe?: string;
    skills?: string[];
    interests?: string[];
    pastProjects?: string[];
    // Index signature for dynamic access
    [key: string]: string | string[] | undefined;
}

export interface IGraphPresence {
    id: string;
    availability: string;
    activity: string;
}

export interface IGraphManager {
    id: string;
    displayName: string;
    jobTitle?: string;
    department?: string;
}

export interface IGraphSkill {
    id: string;
    allowedAudiences: string;
    createdDateTime: string;
    displayName: string;
    proficiency?: string;
    webUrl?: string;
}

export interface IGraphInterest {
    id: string;
    allowedAudiences: string;
    createdDateTime: string;
    displayName: string;
    categories?: string[];
    webUrl?: string;
}


export class GraphService {
    private readonly context: WebPartContext;
    private graphClient: MSGraphClientV3 | null = null;

    constructor(context: WebPartContext) {
        this.context = context;
    }

    private async getGraphClient(): Promise<MSGraphClientV3> {
        this.graphClient ??= await this.context.msGraphClientFactory.getClient('3');
        return this.graphClient;
    }

    /**
     * Get unique values for a specific field across all users (useful for dropdown filters)
     * @param field The field name to get unique values for (e.g., 'department' or 'officeLocation')
     * @returns Array of unique non-empty strings
     */
    public async getUniqueValues(field: string): Promise<string[]> {
        try {
            const client = await this.getGraphClient();
            const uniqueValues = new Set<string>();
            let nextLink: string | null = `/users?$select=${field}&$top=999`;

            while (nextLink) {
                const response = await client.api(nextLink).get();
                const users = response.value || [];
                users.forEach((u: IGraphUser) => {
                    const val = u[field];
                    if (val && typeof val === 'string') uniqueValues.add(val);
                });
                nextLink = response['@odata.nextLink'] || null;

                // Safety break to prevent infinite loops and limit search
                if (uniqueValues.size > 500 || !nextLink) break;
            }

            return Array.from(uniqueValues).sort((a, b) => a.localeCompare(b));
        } catch (error) {
            console.error('[GraphService] Error in getUniqueValues:', error);
            return [];
        }
    }

    /**
     * Get users from the organization with pagination support
     * @param pageSize Number of users to fetch per page
     * @returns Object containing users and a nextLink for more results
     */
    public async getUsers(
        pageSize: number = 10,
        filterLetter?: string,
        filterType?: string,
        filterValue?: string,
        filterSecondaryValue?: string
    ): Promise<{ users: IGraphUser[]; nextLink: string | null }> {
        try {
            const client = await this.getGraphClient();
            let query = client
                .api('/users')
                .select('id,displayName,givenName,surname,jobTitle,department,officeLocation,city,state,country,mail,userPrincipalName,mobilePhone,businessPhones,onPremisesExtensionAttributes')
                .top(pageSize);

            const filters: string[] = [];

            if (filterLetter) {
                filters.push(`startswith(displayName,'${filterLetter}')`);
            }

            if (filterType && filterType !== 'none' && filterValue) {
                switch (filterType) {
                    case 'department':
                        filters.push(`startswith(department,'${filterValue}')`);
                        break;
                    case 'location':
                        filters.push(`startswith(officeLocation,'${filterValue}')`);
                        break;
                    case 'extension':
                        if (filterSecondaryValue) {
                            filters.push(`onPremisesExtensionAttributes/${filterValue} eq '${filterSecondaryValue}'`);
                        }
                        break;
                }
            }

            if (filters.length > 0) {
                query = query.filter(filters.join(' and '));
            }

            const response = await query.get();
            let users: IGraphUser[] = response.value || [];
            const nextLink = response['@odata.nextLink'] || null;

            // Client-side fallback for 'domain' filter
            if (filterType === 'domain' && filterValue) {
                const domainVal = filterValue.toLowerCase();
                users = users.filter(u =>
                    u.mail?.toLowerCase().endsWith(`@${domainVal}`) ||
                    u.userPrincipalName?.toLowerCase().endsWith(`@${domainVal}`)
                );
            }

            return {
                users: users,
                nextLink: nextLink
            };
        } catch (error) {
            console.error('[GraphService] Error in getUsers:', error);
            return { users: [], nextLink: null };
        }
    }

    /**
     * Fetch more users using the nextLink provided by a previous request
     * @param nextLink The URL for the next page of results
     * @returns Object containing users and a nextLink for subsequent results
     */
    public async getMoreUsers(
        nextLink: string,
        filterType?: string,
        filterValue?: string
    ): Promise<{ users: IGraphUser[]; nextLink: string | null }> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api(nextLink)
                .get();

            let users: IGraphUser[] = response.value || [];
            const nextLinkResult = response['@odata.nextLink'] || null;

            // Client-side fallback for 'domain' filter
            if (filterType === 'domain' && filterValue) {
                const domainVal = filterValue.toLowerCase();
                users = users.filter(u =>
                    u.mail?.toLowerCase().endsWith(`@${domainVal}`) ||
                    u.userPrincipalName?.toLowerCase().endsWith(`@${domainVal}`)
                );
            }

            return {
                users: users,
                nextLink: nextLinkResult
            };
        } catch (error) {
            console.error('[GraphService] Error in getMoreUsers:', error);
            return { users: [], nextLink: null };
        }
    }

    /**
     * Get a specific user by ID
     * @param userId User ID
     * @returns User object
     */
    public async getUser(userId: string): Promise<IGraphUser | null> {
        try {
            const client = await this.getGraphClient();
            const user = await client
                .api(`/users/${userId}`)
                .select('id,displayName,givenName,surname,jobTitle,department,officeLocation,city,state,country,mail,userPrincipalName,mobilePhone,businessPhones')
                .get();

            return user;
        } catch (error) {
            console.error('[GraphService] Error in getUser:', error);
            return null;
        }
    }

    /**
     * Get multiple users by their emails (User Principal Names)
     * @param emails Array of emails
     * @returns Array of users
     */
    public async getUsersByEmails(emails: string[]): Promise<IGraphUser[]> {
        if (!emails || emails.length === 0) return [];

        try {
            const client = await this.getGraphClient();

            // Build filter string: mail eq 'email1' or mail eq 'email2' ...
            // Graph has limits on query length, so we might need to batch if emails.length > 15
            const batchSize = 15;
            let results: IGraphUser[] = [];

            for (let i = 0; i < emails.length; i += batchSize) {
                const batch = emails.slice(i, i + batchSize);
                const filter = batch.map(email => `mail eq '${email}' or userPrincipalName eq '${email}'`).join(' or ');

                const response = await client
                    .api('/users')
                    .filter(filter)
                    .select('id,displayName,givenName,surname,jobTitle,department,officeLocation,city,state,country,mail,userPrincipalName,mobilePhone,businessPhones')
                    .get();

                if (response?.value) {
                    results = [...results, ...response.value];
                }
            }

            return results;
        } catch (error) {
            console.error('[GraphService] Error in getUsersByEmails:', error);
            return [];
        }
    }

    /**
     * Get current logged-in user
     * @returns Current user object
     */
    public async getCurrentUser(): Promise<IGraphUser | null> {
        try {
            const client = await this.getGraphClient();
            const user = await client
                .api('/me')
                .select('id,displayName,givenName,surname,jobTitle,department,officeLocation,city,state,country,mail,userPrincipalName,mobilePhone,businessPhones')
                .get();

            return user;
        } catch (error) {
            console.error('[GraphService] Error in getCurrentUser:', error);
            return null;
        }
    }

    /**
     * Get detailed user profile including extended fields
     * @param userId User ID
     * @returns User object with extended profile fields
     */
    public async getUserDetails(userId: string): Promise<IGraphUser | null> {
        try {
            const client = await this.getGraphClient();
            const endpoint = userId === 'me' ? '/me' : `/users/${userId}`;
            const user = await client
                .api(endpoint)
                .select('id,displayName,givenName,surname,jobTitle,department,officeLocation,city,state,country,mail,userPrincipalName,mobilePhone,businessPhones,aboutMe,skills,interests,pastProjects')
                .get();

            return user;
        } catch (error) {
            console.error('[GraphService] Error in getUserDetails:', error);
            // Fall back to basic user info if extended fields fail
            return await this.getUser(userId);
        }
    }

    /**
     * Get user's manager
     * @param userId User ID
     * @returns Manager object
     */
    public async getUserManager(userId: string): Promise<IGraphManager | null> {
        try {
            const client = await this.getGraphClient();
            const manager = await client
                .api(`/users/${userId}/manager`)
                .select('id,displayName,jobTitle,department')
                .get();

            return manager;
        } catch (error) {
            console.error('[GraphService] Error fetching user manager:', error);
            return null;
        }
    }

    /**
     * Get user's direct reports
     * @param userId User ID
     * @returns Array of direct reports
     */
    public async getUserDirectReports(userId: string): Promise<IGraphUser[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api(`/users/${userId}/directReports`)
                .select('id,displayName,givenName,surname,jobTitle,department')
                .get();

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error fetching direct reports:', error);
            return [];
        }
    }

    /**
     * Get user's presence (online status)
     * @param userId User ID
     * @returns Presence object
     */
    public async getUserPresence(userId: string): Promise<IGraphPresence | null> {
        try {
            const client = await this.getGraphClient();
            const presence = await client
                .api(`/users/${userId}/presence`)
                .get();

            return presence;
        } catch (error) {
            console.error('[GraphService] Error fetching user presence:', error);
            return null;
        }
    }

    /**
     * Get user's profile photo
     * @param userId User ID
     * @returns Blob URL for the photo or null
     */
    public async getUserPhoto(userId: string): Promise<string | null> {
        try {
            const client = await this.getGraphClient();
            const photoBlob = await client
                .api(`/users/${userId}/photo/$value`)
                .get();

            // Create a blob URL for the photo
            const photoUrl = URL.createObjectURL(photoBlob);
            return photoUrl;
        } catch (error) {
            // User might not have a photo, which results in a 404. 
            // This is a common case and not a failure of the service logic.
            console.log('[GraphService] No photo found for user:', userId, error);
            return null;
        }
    }

    /**
     * Get presence for multiple users (batch request)
     * @param userIds Array of user IDs
     * @returns Array of presence objects
     */
    public async getBatchPresence(userIds: string[]): Promise<IGraphPresence[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api('/communications/getPresencesByUserId')
                .post({
                    ids: userIds
                });

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error fetching batch presence:', error);
            return [];
        }
    }

    /**
     * Search users by display name
     * @param searchTerm Search term
     * @param top Number of results (default: 50)
     * @returns Array of matching users
     */
    public async searchUsers(searchTerm: string, top: number = 50): Promise<IGraphUser[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api('/users')
                .select('id,displayName,givenName,surname,jobTitle,department,mail,userPrincipalName')
                .filter(`startswith(displayName,'${searchTerm}') or startswith(mail,'${searchTerm}')`)
                .top(top)
                .get();

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error searching users:', error);
            return [];
        }
    }

    /**
     * Get users by department
     * @param department Department name
     * @returns Array of users in the department
     */
    public async getUsersByDepartment(department: string): Promise<IGraphUser[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api('/users')
                .select('id,displayName,givenName,surname,jobTitle,department,mail,userPrincipalName')
                .filter(`department eq '${department}'`)
                .get();

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error fetching users by department:', error);
            return [];
        }
    }

    /**
     * Get user's colleagues (people working with)
     * @param userId User ID
     * @returns Array of colleagues
     */
    public async getUserColleagues(userId: string): Promise<IGraphUser[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api(`/users/${userId}/people`)
                .select('id,displayName,givenName,surname,jobTitle,department')
                .filter("personType/class eq 'Person' and personType/subclass eq 'OrganizationUser'")
                .top(20)
                .get();

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error fetching colleagues:', error);
            return [];
        }
    }

    /**
     * Get user's skills from profile
     * @param userId User ID (use 'me' for current user)
     * @returns Array of skills
     */
    public async getUserSkills(userId: string = 'me'): Promise<IGraphSkill[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api(`/users/${userId}/profile/skills`)
                .get();

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error fetching user skills:', error);
            return [];
        }
    }

    /**
     * Add a skill to user's profile
     * @param userId User ID (use 'me' for current user)
     * @param skillName Skill name
     * @param proficiency Proficiency level (optional)
     * @returns Created skill object
     */
    public async addUserSkill(skillName: string, proficiency?: string, userId: string = 'me'): Promise<IGraphSkill | null> {
        try {
            const client = await this.getGraphClient();
            const skill = await client
                .api(`/users/${userId}/profile/skills`)
                .post({
                    displayName: skillName,
                    proficiency: proficiency,
                    allowedAudiences: 'organization'
                });

            return skill;
        } catch (error) {
            console.error('[GraphService] Error adding user skill:', error);
            return null;
        }
    }

    /**
     * Update a user's skill
     * @param userId User ID (use 'me' for current user)
     * @param skillId Skill ID
     * @param skillName Updated skill name
     * @param proficiency Updated proficiency level
     * @returns Updated skill object
     */
    public async updateUserSkill(skillId: string, skillName: string, proficiency?: string, userId: string = 'me'): Promise<IGraphSkill | null> {
        try {
            const client = await this.getGraphClient();
            const skill = await client
                .api(`/users/${userId}/profile/skills/${skillId}`)
                .patch({
                    displayName: skillName,
                    proficiency: proficiency
                });

            return skill;
        } catch (error) {
            console.error('[GraphService] Error updating user skill:', error);
            return null;
        }
    }

    /**
     * Delete a user's skill
     * @param userId User ID (use 'me' for current user)
     * @param skillId Skill ID
     * @returns Success boolean
     */
    public async deleteUserSkill(skillId: string, userId: string = 'me'): Promise<boolean> {
        try {
            const client = await this.getGraphClient();
            await client
                .api(`/users/${userId}/profile/skills/${skillId}`)
                .delete();

            return true;
        } catch (error) {
            console.error('[GraphService] Error deleting user skill:', error);
            return false;
        }
    }

    /**
     * Get user's interests from profile
     * @param userId User ID (use 'me' for current user)
     * @returns Array of interests
     */
    public async getUserInterests(userId: string = 'me'): Promise<IGraphInterest[]> {
        try {
            const client = await this.getGraphClient();
            const response = await client
                .api(`/users/${userId}/profile/interests`)
                .get();

            return response.value || [];
        } catch (error) {
            console.error('[GraphService] Error fetching user interests:', error);
            return [];
        }
    }

    /**
     * Add an interest to user's profile
     * @param userId User ID (use 'me' for current user)
     * @param interestName Interest name
     * @param categories Categories (optional)
     * @returns Created interest object
     */
    public async addUserInterest(interestName: string, categories?: string[], userId: string = 'me'): Promise<IGraphInterest | null> {
        try {
            const client = await this.getGraphClient();
            const interest = await client
                .api(`/users/${userId}/profile/interests`)
                .post({
                    displayName: interestName,
                    categories: categories,
                    allowedAudiences: 'organization'
                });

            return interest;
        } catch (error) {
            console.error('[GraphService] Error adding user interest:', error);
            return null;
        }
    }

    /**
     * Delete a user's interest
     * @param userId User ID (use 'me' for current user)
     * @param interestId Interest ID
     * @returns Success boolean
     */
    public async deleteUserInterest(interestId: string, userId: string = 'me'): Promise<boolean> {
        try {
            const client = await this.getGraphClient();
            await client
                .api(`/users/${userId}/profile/interests/${interestId}`)
                .delete();

            return true;
        } catch (error) {
            console.error('[GraphService] Error deleting user interest:', error);
            return false;
        }
    }

    /**
     * Update current user's profile fields
     * @param fields Fields to update
     * @returns Object with success status
     */
    public async updateUserProfile(fields: Partial<IGraphUser>): Promise<{ success: boolean; details?: any }> {
        try {
            const client = await this.getGraphClient();

            // Use Graph Beta PATCH /me for profile updates (supports skills and interests directly)
            await client.api('/me').version('beta').patch(fields);

            return { success: true };
        } catch (error) {
            return { success: false, details: (error as Error).message };
        }
    }
}
