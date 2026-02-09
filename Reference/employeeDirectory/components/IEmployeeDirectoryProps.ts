import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEmployeeDirectoryProps {
    description: string;
    excludeUsers: string;
    filterFields: string[];
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
    cardFields: string[];
    context: WebPartContext;
    filterType?: 'none' | 'domain' | 'extension' | 'department' | 'location';
    filterValue?: string;
    filterSecondaryValue?: string;
    title?: string;
    enableStarred?: boolean;
    starredUsers?: any[];
    titleFontSize?: number;
    descriptionFontSize?: number;
    subHeadingFontSize?: number;
    contentFontSize?: number;
    headerPadding?: number;        // For Main Directory Header
    profileHeaderPadding?: number; // For Profile View Header
    profileHeaderHeight?: number;  // Explicit height for profile header bar
    enableSelfUpdate?: boolean;    // Enable self-service profile update
    editFields?: string[];         // Fields allowed for editing
    headerAlignment?: 'left' | 'center' | 'right';
    themeTemplate?: 'default' | 'berry' | 'orange' | 'grape' | 'forest' | 'ocean' | 'midnight' | 'mint';
    visualTheme?: 'modern-scrolling-premium' | 'modern-tabbed-premium';
    kudosListId?: string;
    kudosTargetField?: string;
    kudosMessageField?: string;
    kudosTypeField?: string;
    kudosDateField?: string;
    enableKudos?: boolean;
}
