import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IModernEmployeeDirectoryProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  profileLayout: 'scroll' | 'tab';
  mainHeadingSize: number;
  subHeadingSize: number;
  contentHeadingSize: number;
  context: WebPartContext;
  // Kudos System
  enableKudos: boolean;
  kudosListId: string;
  kudosRecipientColumn: string;
  kudosAuthorColumn: string;
  kudosMessageColumn: string;
  kudosBadgeTypeColumn: string;
  orgChartLayout: 'vertical' | 'horizontal' | 'compact';
  updatableFields: string[];
  enablePagination: boolean;
  pageSize: number;
  paginationType: 'loadMore' | 'prevNext';
  // Hall of Fame
  featuredEmails: string[];
  minKudosCount: number;
  // Filters
  filterType?: 'none' | 'domain' | 'extension' | 'department' | 'location';
  filterValue?: string;
  filterSecondaryValue?: string;
  homePageFilterFields: string[];
  dynamicFilterData?: { [key: string]: string[] };
  // Audit Logging
  enableAudit: boolean;
  auditListId: string;
  auditActivityColumn: string;
  auditActorColumn: string;
  auditTargetColumn: string;
  auditDetailsColumn: string;
  enableAuditDebug: boolean;
  containerMargin: number;
  badgeCircleSize: number;
  badgeFontSize: number;

}

