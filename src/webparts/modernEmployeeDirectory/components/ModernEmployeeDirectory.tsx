import * as React from 'react';
import styles from './ModernEmployeeDirectory.module.scss';
import type { IModernEmployeeDirectoryProps } from './IModernEmployeeDirectoryProps';
import { TopBar } from './Shared/TopBar';
import { SidePanel } from './Shared/SidePanel';
import { DirectoryHome, IEmployee } from './Home/DirectoryHome';
import { ProfileScrolling } from './Profile/ProfileScrolling';
import { ProfileTabbed } from './Profile/ProfileTabbed';
import { KudosService } from '../services/KudosService';
import { GraphService, IGraphUser } from '../services/GraphService';
import { AuditService } from '../services/AuditService';
import { IKudos } from '../models/IKudos';

export type SidePanelMode = 'profile' | 'kudos' | 'updateProfile';

export interface IModernEmployeeDirectoryState {
  currentView: 'HOME' | 'PROFILE_SCROLL' | 'PROFILE_TAB';
  selectedEmployee: IEmployee | null;
  isSidePanelOpen: boolean;
  sidePanelMode: SidePanelMode;
  kudos: IKudos[];
  kudosLoading: boolean;
  currentUser: {
    id: string;
    initials: string;
    name: string;
    email: string;
    photoUrl?: string;
    jobTitle?: string;
    aboutMe?: string;
    mobilePhone?: string;
    officeLocation?: string;
    skills?: string[];
    interests?: string[];
    responsibilities?: string[];
    projects?: string[];
    pastProjects?: string[];
  };
  employees: IEmployee[];
  employeesLoading: boolean;
  selectedEmployeeKudosCount: number;
  badgeChoices: string[];
  currentUserLoading: boolean;
  topKudosEmails: string[];
  nextPageLink: string | null;
  prevPageLinks: string[];
  lastFetchedLink: string | null;
  selectedLetter: string | null;
  dynamicFilterData: { [key: string]: string[] };
  auditLogs: Array<{ message: string, type: 'info' | 'success' | 'error' | 'warning', timestamp: Date }>;
  showDebugPanel: boolean;
}

const MOCK_CURRENT_USER = {
  initials: 'CU',
  name: 'Current User',
  email: 'user@modern.com'
};

export default class ModernEmployeeDirectory extends React.Component<IModernEmployeeDirectoryProps, IModernEmployeeDirectoryState> {
  private kudosService: KudosService | null = null;
  private readonly graphService: GraphService | null = null;
  private auditService: AuditService | null = null;
  private _loadToken: number = 0;

  constructor(props: IModernEmployeeDirectoryProps) {
    super(props);
    this.state = {
      currentView: 'HOME',
      selectedEmployee: null,
      isSidePanelOpen: false,
      sidePanelMode: 'profile',
      kudos: [],
      kudosLoading: false,
      currentUser: { ...MOCK_CURRENT_USER, id: 'me' },
      employees: [],
      employeesLoading: false,
      selectedEmployeeKudosCount: 0,
      badgeChoices: [],
      currentUserLoading: false,
      topKudosEmails: [],
      nextPageLink: null,
      prevPageLinks: [],
      lastFetchedLink: null,
      selectedLetter: 'STAR',
      dynamicFilterData: {},
      auditLogs: [],
      showDebugPanel: false
    };

    if (props.context) {
      this.graphService = new GraphService(props.context);
    }

    if (props.enableKudos && props.kudosListId && props.context) {
      this.kudosService = new KudosService(
        props.context,
        {
          listId: props.kudosListId,
          recipientColumn: props.kudosRecipientColumn,
          authorColumn: props.kudosAuthorColumn,
          messageColumn: props.kudosMessageColumn,
          badgeTypeColumn: props.kudosBadgeTypeColumn
        }
      );
    }

    if (props.enableAudit && props.auditListId) {
      this.auditService = new AuditService(props.context, {
        listId: props.auditListId,
        activityColumn: props.auditActivityColumn,
        actorColumn: props.auditActorColumn,
        targetColumn: props.auditTargetColumn,
        detailsColumn: props.auditDetailsColumn,
        onLog: (msg, type) => this._handleAuditLogMessage(msg, type)
      });
    }
  }

  private _handleAuditLogMessage(message: string, type: 'info' | 'success' | 'error' | 'warning'): void {
    this.setState(prevState => ({
      auditLogs: [{ message, type, timestamp: new Date() }, ...prevState.auditLogs.slice(0, 49)]
    }));
  }

  private readonly _fetchDynamicFilterValues = async (): Promise<void> => {
    const { homePageFilterFields } = this.props;
    if (!homePageFilterFields || homePageFilterFields.length === 0 || !this.graphService) return;

    const dynamicFilterData: { [key: string]: string[] } = {};

    // Fetch unique values for each configured field in parallel
    const fetchPromises = homePageFilterFields.map(async field => {
      try {
        const values = await this.graphService!.getUniqueValues(field);
        dynamicFilterData[field] = values;
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching dynamic filter values for field:', field, error);
        dynamicFilterData[field] = [];
      }
    });

    await Promise.all(fetchPromises);
    this.setState({ dynamicFilterData });
  }

  private async _fetchTopKudosEmails(): Promise<void> {
    if (this.kudosService && this.props.minKudosCount > 0) {
      try {
        const topKudosEmails = await this.kudosService.getHallOfFameCandidates(this.props.minKudosCount);
        this.setState({ topKudosEmails });
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching top kudos emails:', error);
      }
    }
  }

  public componentDidMount(): void {
    (async (): Promise<void> => {
      this._fetchTopKudosEmails().catch(err => { console.error(err); });

      try {
        const currentUser = this.props.context.pageContext.user;

        this.setState({
          currentUser: {
            id: currentUser.loginName || currentUser.email,
            initials: this._getInitials(currentUser.displayName),
            name: currentUser.displayName,
            email: currentUser.loginName || currentUser.email
          }
        }, () => {
          this._initializeCurrentUser().catch(err => { console.error(err); });
        });
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error in componentDidMount:', error);
      }

      await this._loadEmployees();
      await this._fetchDynamicFilterValues();

      if (this.props.enableKudos && this.kudosService) {
        try {
          const choices = await this.kudosService.getBadgeTypeChoices();
          this.setState({ badgeChoices: choices });
        } catch (error) {
          console.error('[ModernEmployeeDirectory] Error loading badge choices:', error);
        }
      }
    })().catch(err => { console.error(err); });
  }

  private async _initializeCurrentUser(): Promise<void> {
    if (!this.graphService) return;

    try {
      const currentUser = await this.graphService.getCurrentUser();

      if (currentUser) {
        let currentUserId = currentUser.id;

        this.setState(prevState => ({
          currentUser: {
            ...prevState.currentUser,
            id: currentUser.id,
            jobTitle: currentUser.jobTitle,
            mobilePhone: currentUser.mobilePhone,
            aboutMe: currentUser.aboutMe,
            officeLocation: currentUser.officeLocation,
            skills: currentUser.skills,
            interests: currentUser.interests,
            pastProjects: currentUser.pastProjects
          }
        }));

        // If ID is missing (rare), try to get it from context
        if (!currentUserId) {
          try {
            const me = await this.graphService.getUser(this.props.context.pageContext.user.email);
            if (me) {
              currentUserId = me.id;
              this.setState(prevState => ({
                currentUser: { ...prevState.currentUser, id: me.id }
              }));
            }
          } catch (error) {
            console.error('[ModernEmployeeDirectory] Error fetching current user ID from context:', error);
          }
        }

        // Fetch photo
        if (currentUserId) {
          const photo = await this.graphService.getUserPhoto(currentUserId);
          if (photo) {
            this.setState(prevState => ({
              currentUser: { ...prevState.currentUser, photoUrl: photo }
            }));
          }
        }
      }
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error fetching current user:', error);
    }
  }

  public componentDidUpdate(prevProps: IModernEmployeeDirectoryProps): void {
    (async (): Promise<void> => {
      // Handle Filter Field Changes
      if (JSON.stringify(prevProps.homePageFilterFields) !== JSON.stringify(this.props.homePageFilterFields)) {
        await this._fetchDynamicFilterValues();
      }

      // Handle Layout Changes
      if (prevProps.profileLayout !== this.props.profileLayout && this.state.selectedEmployee) {
        this.setState({
          currentView: this.props.profileLayout === 'scroll' ? 'PROFILE_SCROLL' : 'PROFILE_TAB'
        });
      }

      // Handle Component-Specific Setting Changes
      await this._handleKudosPropChanges(prevProps);
      await this._handleFilterSettingChanges(prevProps);
      this._handleAuditPropChanges(prevProps);
    })().catch(err => { console.error(err); });
  }

  private async _handleKudosPropChanges(prevProps: IModernEmployeeDirectoryProps): Promise<void> {
    const kudosPropsChanged = prevProps.enableKudos !== this.props.enableKudos ||
      prevProps.kudosListId !== this.props.kudosListId ||
      prevProps.kudosRecipientColumn !== this.props.kudosRecipientColumn ||
      prevProps.kudosAuthorColumn !== this.props.kudosAuthorColumn ||
      prevProps.kudosMessageColumn !== this.props.kudosMessageColumn ||
      prevProps.kudosBadgeTypeColumn !== this.props.kudosBadgeTypeColumn;

    if (kudosPropsChanged) {
      if (this.props.enableKudos && this.props.kudosListId) {
        this.kudosService = new KudosService(this.props.context, {
          listId: this.props.kudosListId,
          recipientColumn: this.props.kudosRecipientColumn,
          authorColumn: this.props.kudosAuthorColumn,
          messageColumn: this.props.kudosMessageColumn,
          badgeTypeColumn: this.props.kudosBadgeTypeColumn
        });

        try {
          const choices = await this.kudosService.getBadgeTypeChoices();
          this.setState({ badgeChoices: choices });
        } catch (error) {
          console.error('[ModernEmployeeDirectory] Error loading badge choices:', error);
        }
      } else {
        this.kudosService = null;
      }
    }
  }

  private async _handleFilterSettingChanges(prevProps: IModernEmployeeDirectoryProps): Promise<void> {
    const filterSettingChanged = prevProps.filterType !== this.props.filterType ||
      prevProps.filterValue !== this.props.filterValue ||
      prevProps.filterSecondaryValue !== this.props.filterSecondaryValue;

    if (filterSettingChanged) {
      await this._loadEmployees();
    }
  }

  private _handleAuditPropChanges(prevProps: IModernEmployeeDirectoryProps): void {
    const auditPropsChanged = prevProps.enableAudit !== this.props.enableAudit ||
      prevProps.auditListId !== this.props.auditListId ||
      prevProps.auditActivityColumn !== this.props.auditActivityColumn ||
      prevProps.auditActorColumn !== this.props.auditActorColumn ||
      prevProps.auditTargetColumn !== this.props.auditTargetColumn ||
      prevProps.auditDetailsColumn !== this.props.auditDetailsColumn;

    if (auditPropsChanged) {
      if (this.props.enableAudit && this.props.auditListId) {
        this.auditService = new AuditService(this.props.context, {
          listId: this.props.auditListId,
          activityColumn: this.props.auditActivityColumn,
          actorColumn: this.props.auditActorColumn,
          targetColumn: this.props.auditTargetColumn,
          detailsColumn: this.props.auditDetailsColumn,
          onLog: (msg, type) => this._handleAuditLogMessage(msg, type)
        });
      } else {
        this.auditService = null;
      }
    }
  }

  private async _loadEmployees(): Promise<void> {
    const token = ++this._loadToken;
    this.setState({ employeesLoading: true });

    try {
      if (this.graphService) {
        let employees: IEmployee[] = [];
        let nextLink: string | null = null;

        if (this.state.selectedLetter === 'STAR') {
          employees = await this._loadHallOfFame();
        } else {
          const pageSize = this.props.enablePagination ? this.props.pageSize : 999;
          const response = await this.graphService.getUsers(
            pageSize,
            this.state.selectedLetter || undefined,
            this.props.filterType,
            this.props.filterValue,
            this.props.filterSecondaryValue
          );
          employees = this._mapGraphUsersToEmployees(response.users);
          nextLink = response.nextLink;
        }

        if (token !== this._loadToken) return;

        this.setState({
          employees,
          employeesLoading: false,
          nextPageLink: nextLink,
          prevPageLinks: []
        });

        this._fetchEmployeePhotos(employees, token).catch(err => { console.error(err); });
        this._fetchManagerInfo(employees, token).catch(err => { console.error(err); });
      }
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error in _loadEmployees:', error);
      this.setState({ employeesLoading: false });
    }
  }

  private readonly _handleLoadMore = async (): Promise<void> => {
    const { nextPageLink } = this.state;
    if (!nextPageLink || !this.graphService) return;
    this._handleAuditLog('Pagination', 'Load More', { functionName: '_handleLoadMore' });

    const token = ++this._loadToken;
    this.setState({ employeesLoading: true });

    try {
      const response = await this.graphService.getMoreUsers(
        nextPageLink,
        this.props.filterType,
        this.props.filterValue
      );
      const newEmployees = this._mapGraphUsersToEmployees(response.users);

      if (token !== this._loadToken) return;

      this.setState(prevState => {
        const combined = [...prevState.employees, ...newEmployees];
        const unique = combined.filter((emp, index, self) =>
          index === self.findIndex((e) => e.id === emp.id)
        );
        return {
          employees: unique,
          employeesLoading: false,
          nextPageLink: response.nextLink
        };
      });

      this._fetchEmployeePhotos(newEmployees, token).catch(err => { console.error(err); });
      this._fetchManagerInfo(newEmployees, token).catch(err => { console.error(err); });
    } catch {
      this.setState({ employeesLoading: false });
    }
  }

  private readonly _handleNextPage = async (): Promise<void> => {
    const { nextPageLink, prevPageLinks, lastFetchedLink } = this.state;
    if (!nextPageLink || !this.graphService) return;

    const token = ++this._loadToken;
    this.setState({ employeesLoading: true });

    try {
      const response = await this.graphService.getMoreUsers(
        nextPageLink,
        this.props.filterType,
        this.props.filterValue
      );
      const newEmployees = this._mapGraphUsersToEmployees(response.users);

      if (token !== this._loadToken) return;

      const linkToStore = prevPageLinks.length === 0 ? 'INITIAL' : lastFetchedLink;

      this.setState(prevState => ({
        employees: newEmployees,
        employeesLoading: false,
        nextPageLink: response.nextLink,
        prevPageLinks: [...prevState.prevPageLinks, linkToStore!],
        lastFetchedLink: nextPageLink
      }));

      this._fetchEmployeePhotos(newEmployees, token).catch(err => { console.error(err); });
      this._fetchManagerInfo(newEmployees, token).catch(err => { console.error(err); });
    } catch {
      this.setState({ employeesLoading: false });
    }
  }

  private readonly _handlePrevPage = async (): Promise<void> => {
    const { prevPageLinks } = this.state;
    if (prevPageLinks.length === 0 || !this.graphService) return;

    const token = ++this._loadToken;
    const newStack = [...prevPageLinks];
    const prevLink = newStack.pop();

    this.setState({ employeesLoading: true, prevPageLinks: newStack });

    try {
      let response;
      if (prevLink === 'INITIAL') {
        const pageSize = this.props.enablePagination ? this.props.pageSize : 999;
        response = await this.graphService.getUsers(
          pageSize,
          this.state.selectedLetter || undefined,
          this.props.filterType,
          this.props.filterValue,
          this.props.filterSecondaryValue
        );
      } else {
        response = await this.graphService.getMoreUsers(
          prevLink!,
          this.props.filterType,
          this.props.filterValue
        );
      }

      if (token !== this._loadToken) return;

      const employees = this._mapGraphUsersToEmployees(response.users);
      this.setState({
        employees,
        employeesLoading: false,
        nextPageLink: response.nextLink
      });

      this._fetchEmployeePhotos(employees, token).catch(err => { console.error(err); });
      this._fetchManagerInfo(employees, token).catch(err => { console.error(err); });
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error in _handlePrevPage:', error);
      this.setState({ employeesLoading: false });
    }
  }

  private readonly _handleLetterChange = async (letter: string | null): Promise<void> => {
    this._handleAuditLog('Navigation', `Alphabet Filter: ${letter || 'All'}`, { functionName: '_handleLetterChange', letter });
    this.setState({
      selectedLetter: letter,
      prevPageLinks: [],
      nextPageLink: null,
      lastFetchedLink: null
    }, async () => {
      await this._loadEmployees();
    });
  }

  private async _fetchEmployeePhotos(employees: IEmployee[], token: number): Promise<void> {
    if (!this.graphService) return;

    const batchSize = 10;
    for (let i = 0; i < employees.length; i += batchSize) {
      if (token !== this._loadToken) return;
      const batch = employees.slice(i, i + batchSize);

      await Promise.all(batch.map(async (emp) => {
        try {
          const photoUrl = await this.graphService!.getUserPhoto(emp.id);
          if (photoUrl && token === this._loadToken) {
            this.setState(prevState => ({
              employees: prevState.employees.map(e =>
                e.id === emp.id ? { ...e, photoUrl } : e
              )
            }));
          }
        } catch (error) {
          console.error('[ModernEmployeeDirectory] Error fetching photo for employee:', emp.id, error);
        }
      }));
    }
  }

  private async _fetchManagerInfo(employees: IEmployee[], token: number): Promise<void> {
    const graphService = this.graphService;
    if (!graphService) return;

    const batchSize = 10;
    for (let i = 0; i < employees.length; i += batchSize) {
      if (token !== this._loadToken) return;
      const batch = employees.slice(i, i + batchSize);

      await Promise.all(batch.map(async (emp) => {
        try {
          const manager = await graphService.getUserManager(emp.id);
          if (manager && token === this._loadToken) {
            this.setState(prevState => {
              const managerExists = prevState.employees.some(e => e.id === manager.id);
              let updatedEmployees = prevState.employees.map(e =>
                e.id === emp.id ? { ...e, managerId: manager.id } : e
              );

              if (!managerExists) {
                const managerEmp: IEmployee = {
                  id: manager.id,
                  displayName: manager.displayName,
                  initials: this._getInitials(manager.displayName),
                  jobTitle: manager.jobTitle,
                  department: manager.department,
                  isOnline: false,
                  _isOrgChartNode: true
                };
                updatedEmployees = [...updatedEmployees, managerEmp];
              }

              return { employees: updatedEmployees };
            });
          }
        } catch (error) {
          console.error('[ModernEmployeeDirectory] Error fetching manager for employee:', emp.id, error);
        }
      }));
    }
  }

  private _mapGraphUsersToEmployees(graphUsers: IGraphUser[]): IEmployee[] {
    return graphUsers.map(user => ({
      id: user.id,
      displayName: user.displayName,
      givenName: user.givenName,
      surname: user.surname,
      jobTitle: user.jobTitle,
      department: user.department,
      officeLocation: user.officeLocation,
      city: user.city,
      state: user.state,
      country: user.country,
      mail: user.mail,
      userPrincipalName: user.userPrincipalName,
      mobilePhone: user.mobilePhone,
      businessPhones: user.businessPhones,
      initials: this._getInitials(user.displayName),
      isOnline: false,
      isFeatured: (this.props.featuredEmails || []).some((email: string) => email.toLowerCase() === user.mail?.toLowerCase()),
      isTopKudos: this.state.topKudosEmails.some((email: string) => email.toLowerCase() === user.mail?.toLowerCase())
    }));
  }

  private _getInitials(name: string): string {
    if (!name) return 'CU';
    const parts = name.split(' ');
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return name.substring(0, 2).toUpperCase();
  }

  private async _loadHallOfFame(): Promise<IEmployee[]> {
    if (!this.graphService) return [];

    let hofEmails: string[] = [];
    if (this.kudosService && this.props.minKudosCount > 0) {
      try {
        hofEmails = await this.kudosService.getHallOfFameCandidates(this.props.minKudosCount);
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error loading Hall of Fame candidates:', error);
      }
    }

    const featuredEmails = this.props.featuredEmails || [];
    const combinedEmails = Array.from(new Set([...hofEmails, ...featuredEmails]))
      .filter(email => !!email && email.trim().length > 0);

    if (combinedEmails.length === 0) return [];

    try {
      const hofUsers = await this.graphService.getUsersByEmails(combinedEmails);
      return this._mapGraphUsersToEmployees(hofUsers)
        .map(emp => ({
          ...emp,
          isFeatured: featuredEmails.some(fe => fe.toLowerCase() === emp.mail?.toLowerCase()),
          isTopKudos: hofEmails.some(he => he.toLowerCase() === emp.mail?.toLowerCase())
        }))
        .filter(emp => emp.isFeatured || emp.isTopKudos);
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error loading Hall of Fame details:', error);
      return [];
    }
  }



  private readonly _handleSelectEmployee = async (employee: IEmployee): Promise<void> => {
    let detailedEmployee = employee;
    if (this.graphService) {
      try {
        const detailedProfile = await this.graphService.getUserDetails(employee.id);
        if (detailedProfile) {
          detailedEmployee = {
            ...employee,
            aboutMe: detailedProfile.aboutMe,
            skills: detailedProfile.skills,
            interests: detailedProfile.interests,
            pastProjects: detailedProfile.pastProjects
          };
        }

        const photoUrl = await this.graphService.getUserPhoto(employee.id);
        if (photoUrl) {
          detailedEmployee.photoUrl = photoUrl;
        }
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching detailed profile for:', employee.id, error);
      }
    }

    let kudosCount = 0;
    if (this.props.enableKudos && this.kudosService && employee.mail) {
      try {
        kudosCount = await this.kudosService.getKudosCount(employee.mail);
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching kudos count for employee:', employee.mail, error);
      }
    }

    this.setState({
      selectedEmployee: detailedEmployee,
      currentView: this.props.profileLayout === 'scroll' ? 'PROFILE_SCROLL' : 'PROFILE_TAB',
      selectedEmployeeKudosCount: kudosCount
    });

    // Log Activity: Profile View
    if (this.auditService) {
      this.auditService.logActivity(
        'Profile View',
        employee.displayName || employee.mail || employee.id,
        JSON.stringify({
          functionName: '_handleSelectEmployee',
          employeeId: employee.id,
          employeeName: employee.displayName,
          employeeMail: employee.mail,
          view: this.props.profileLayout
        }),
        this.state.currentUser.email
      ).catch(err => { console.error(err); });
    }
  }

  private readonly _handleAuditLog = (activity: string, target: string, details: unknown): void => {
    if (this.auditService) {
      this.auditService.logActivity(
        activity,
        target,
        JSON.stringify(details),
        this.state.currentUser.email
      ).catch(err => { console.error(err); });
    }
  }

  private readonly _handleKudosForEmployee = async (employee: IEmployee): Promise<void> => {
    let kudosCount = 0;
    if (this.props.enableKudos && this.kudosService && employee.mail) {
      try {
        kudosCount = await this.kudosService.getKudosCount(employee.mail);
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching kudos count for employee:', employee.mail, error);
      }
    }

    this.setState({
      selectedEmployee: employee,
      selectedEmployeeKudosCount: kudosCount
    }, () => {
      this._toggleSidePanel('kudos');
    });
  }

  private readonly _handleBackToDirectory = (): void => {
    this.setState({
      currentView: 'HOME',
      selectedEmployee: null
    });

    // Activity logging removed per user request
  }

  private readonly _toggleSidePanel = (mode: SidePanelMode = 'profile'): void => {
    const isOpeningDifferentMode = this.state.isSidePanelOpen && this.state.sidePanelMode !== mode;

    if (this.state.isSidePanelOpen && !isOpeningDifferentMode) {
      this.setState({ isSidePanelOpen: false });
      return;
    }

    this.setState({
      isSidePanelOpen: true,
      sidePanelMode: mode
    });

    this._handleAuditLog('Panel Access', mode, { functionName: '_toggleSidePanel', panelMode: mode });

    if (mode === 'updateProfile') {
      this._loadFreshProfileData().catch(err => { console.error(err); });
    }

    if (mode === 'kudos' && this.props.enableKudos) {
      const userId = this.state.selectedEmployee?.mail || this.state.currentUser.email;
      this._loadKudos(userId).catch(err => { console.error(err); });
    }
  }

  private async _loadFreshProfileData(): Promise<void> {
    if (!this.graphService) return;

    try {
      this.setState({ currentUserLoading: true });
      const currentUser = await this.graphService.getCurrentUser();
      const detailedProfile = await this.graphService.getUserDetails('me');

      if (currentUser && detailedProfile) {
        this.setState(prevState => ({
          currentUser: {
            ...prevState.currentUser,
            name: currentUser.displayName,
            jobTitle: currentUser.jobTitle,
            aboutMe: detailedProfile.aboutMe,
            mobilePhone: currentUser.mobilePhone,
            officeLocation: currentUser.officeLocation,
            skills: detailedProfile.skills,
            interests: detailedProfile.interests,
            pastProjects: detailedProfile.pastProjects
          },
          currentUserLoading: false
        }));
      } else {
        this.setState({ currentUserLoading: false });
      }
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error in _loadFreshProfileData:', error);
      this.setState({ currentUserLoading: false });
    }
  }

  private readonly _closeSidePanel = (): void => {
    this.setState({ isSidePanelOpen: false });
  }

  private readonly _loadKudos = async (userId: string): Promise<void> => {
    if (!this.props.enableKudos || !this.kudosService) return;

    this.setState({ kudosLoading: true });
    try {
      const kudos = await this.kudosService.getKudosForUser(userId);
      this.setState({ kudos, kudosLoading: false });
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error in _loadKudos for user:', userId, error);
      this.setState({ kudosLoading: false });
    }
  }

  private readonly _handleGiveKudos = async (message: string, badgeType: string): Promise<void> => {
    if (!this.kudosService || !this.state.selectedEmployee) return;

    const recipient = this.state.selectedEmployee;
    if (!recipient.mail) return;

    try {
      const success = await this.kudosService.giveKudos(
        recipient.mail,
        recipient.displayName,
        message,
        badgeType
      );

      if (success) {
        this._loadKudos(recipient.mail).catch(err => { console.error(err); });

        // Log Activity: Kudos Given
        if (this.auditService) {
          this.auditService.logActivity(
            'Kudos Given',
            recipient.displayName || recipient.mail,
            JSON.stringify({
              functionName: '_handleGiveKudos',
              recipientMail: recipient.mail,
              recipientName: recipient.displayName,
              badgeType: badgeType,
              message: message
            }),
            this.state.currentUser.email
          ).catch(err => { console.error(err); });
        }
      }
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error in _handleGiveKudos:', error);
    }
  }

  private readonly _handleUpdateProfile = async (updatedFields: unknown): Promise<boolean> => {
    if (!this.graphService) return false;

    try {
      const result = await this.graphService.updateUserProfile(updatedFields as IGraphUser);
      if (result.success) {
        const currentUser = await this.graphService.getCurrentUser();
        const detailedProfile = await this.graphService.getUserDetails('me');

        if (currentUser && detailedProfile) {
          this.setState(prevState => ({
            currentUser: {
              ...prevState.currentUser,
              name: currentUser.displayName,
              jobTitle: currentUser.jobTitle,
              aboutMe: detailedProfile.aboutMe,
              mobilePhone: currentUser.mobilePhone,
              officeLocation: currentUser.officeLocation,
              skills: detailedProfile.skills,
              interests: detailedProfile.interests
            }
          }));
        }

        // Log Activity: Profile Update
        if (this.auditService) {
          this.auditService.logActivity(
            'Profile Update',
            this.state.currentUser.name || this.state.currentUser.email,
            JSON.stringify({
              functionName: '_handleUpdateProfile',
              updatedFields: Object.keys(updatedFields as object),
              timestamp: new Date().toISOString()
            }),
            this.state.currentUser.email
          ).catch(err => { console.error(err); });
        }
      }
      return result.success;
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error in _handleUpdateProfile:', error);
      return false;
    }
  }

  public render(): React.ReactElement<IModernEmployeeDirectoryProps> {
    const { hasTeamsContext, contentHeadingSize } = this.props;
    const { currentView, selectedEmployee, isSidePanelOpen } = this.state;

    const fontSizeStyles = {
      '--content-heading-size': `${contentHeadingSize}px`
    } as React.CSSProperties;

    const containerStyle: React.CSSProperties = {
      ...fontSizeStyles,
      margin: `${this.props.containerMargin}px`
    };

    return (
      <section className={`${styles.modernEmployeeDirectory} ${hasTeamsContext ? styles.teams : ''}`} style={containerStyle} >
        {currentView === 'HOME' && (
          <TopBar
            onOpenPanel={(mode: SidePanelMode) => {
              if (mode === 'profile') {
                const myEmployee: IEmployee = {
                  id: this.state.currentUser.id,
                  displayName: this.state.currentUser.name,
                  mail: this.state.currentUser.email,
                  initials: this.state.currentUser.initials,
                  photoUrl: this.state.currentUser.photoUrl,
                  jobTitle: this.state.currentUser.jobTitle,
                  aboutMe: this.state.currentUser.aboutMe,
                  mobilePhone: this.state.currentUser.mobilePhone,
                  officeLocation: this.state.currentUser.officeLocation,
                  skills: this.state.currentUser.skills,
                  interests: this.state.currentUser.interests,
                  responsibilities: this.state.currentUser.responsibilities,
                  projects: this.state.currentUser.projects,
                  pastProjects: this.state.currentUser.pastProjects
                };
                this._handleSelectEmployee(myEmployee).catch(err => { console.error(err); });
                this._handleAuditLog('Menu Interaction', 'My Profile', { functionName: 'TopBar.onOpenPanel', source: 'HamburgerMenu' });
              } else {
                const myEmployee: IEmployee = {
                  id: this.state.currentUser.id,
                  displayName: this.state.currentUser.name,
                  mail: this.state.currentUser.email,
                  initials: this.state.currentUser.initials,
                  photoUrl: this.state.currentUser.photoUrl,
                  jobTitle: this.state.currentUser.jobTitle,
                  aboutMe: this.state.currentUser.aboutMe,
                  mobilePhone: this.state.currentUser.mobilePhone,
                  officeLocation: this.state.currentUser.officeLocation,
                  skills: this.state.currentUser.skills,
                  interests: this.state.currentUser.interests,
                  responsibilities: this.state.currentUser.responsibilities,
                  projects: this.state.currentUser.projects,
                  pastProjects: this.state.currentUser.pastProjects
                };

                if (mode === 'kudos') {
                  this.setState({ selectedEmployee: myEmployee }, () => {
                    this._toggleSidePanel('kudos');
                  });
                } else {
                  this._toggleSidePanel(mode);
                }
              }
            }}
            currentUserInitials={this.state.currentUser.initials}
            buildNumber="0.0.1"
          />
        )}

        <SidePanel
          isOpen={isSidePanelOpen}
          onClose={this._closeSidePanel}
          panelMode={this.state.sidePanelMode}
          onSwitchToKudos={() => this.setState({ sidePanelMode: 'kudos' })}
          onSwitchToUpdateProfile={() => this._toggleSidePanel('updateProfile')}
          user={this.state.currentUser}
          enableKudos={this.props.enableKudos}
          selectedEmployee={selectedEmployee}
          kudos={this.state.kudos}
          kudosLoading={this.state.kudosLoading}
          badgeChoices={this.state.badgeChoices}
          onGiveKudos={this._handleGiveKudos}
          updatableFields={this.props.updatableFields}
          onUpdateProfile={this._handleUpdateProfile}
          loading={this.state.currentUserLoading}
        />

        <div className={styles.contentArea}>
          {currentView === 'HOME' && (
            <DirectoryHome
              employees={this.state.employees}
              onSelectEmployee={this._handleSelectEmployee}
              onKudosForEmployee={this._handleKudosForEmployee}
              enableKudos={this.props.enableKudos}
              onLoadMore={this._handleLoadMore}
              hasMore={!!this.state.nextPageLink}
              loading={this.state.employeesLoading}
              paginationType={this.props.paginationType}
              onNextPage={this._handleNextPage}
              onPrevPage={this._handlePrevPage}
              hasPrev={this.state.prevPageLinks.length > 0}
              selectedLetter={this.state.selectedLetter}
              onLetterChange={this._handleLetterChange}
              homePageFilterFields={this.props.homePageFilterFields}
              dynamicFilterData={this.state.dynamicFilterData}
              onAuditLog={this._handleAuditLog}
              badgeCircleSize={this.props.badgeCircleSize}
              badgeFontSize={this.props.badgeFontSize}

            />
          )}

          {currentView === 'PROFILE_SCROLL' && selectedEmployee && (
            <ProfileScrolling
              employee={selectedEmployee}
              employees={this.state.employees}
              kudosCount={this.state.selectedEmployeeKudosCount}
              onBack={this._handleBackToDirectory}
              onKudosClick={() => this._toggleSidePanel('kudos')}
              onEmployeeSelect={this._handleSelectEmployee}
              orgChartLayout={this.props.orgChartLayout}
              onAuditLog={this._handleAuditLog}
            />
          )}

          {currentView === 'PROFILE_TAB' && selectedEmployee && (
            <ProfileTabbed
              employee={selectedEmployee}
              employees={this.state.employees}
              kudosCount={this.state.selectedEmployeeKudosCount}
              onBack={this._handleBackToDirectory}
              onKudosClick={() => this._toggleSidePanel('kudos')}
              onEmployeeSelect={this._handleSelectEmployee}
              orgChartLayout={this.props.orgChartLayout}
              onAuditLog={this._handleAuditLog}
            />
          )}
        </div>

        {/* Audit Debug Panel */}
        {this.props.enableAuditDebug && (
          <div className={`${styles.debugPanel} ${this.state.showDebugPanel ? '' : styles.collapsed}`}>
            <div className={styles.debugHeader}>
              <button
                type="button"
                className={styles.headerLeft}
                style={{
                  background: 'none',
                  border: 'none',
                  padding: 0,
                  margin: 0,
                  font: 'inherit',
                  color: 'inherit',
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  flexGrow: 1,
                  textAlign: 'left'
                }}
                onClick={() => this.setState(prevState => ({ showDebugPanel: !prevState.showDebugPanel }))}
                aria-expanded={this.state.showDebugPanel}
              >
                <h3>Audit Debug Logs ({this.state.auditLogs.length})</h3>
                <span className={styles.toggleBtn}>{this.state.showDebugPanel ? '▼ Hide' : '▲ Show'}</span>
              </button>
              {this.state.showDebugPanel && (
                <button
                  type="button"
                  className={styles.clearBtn}
                  onClick={() => this.setState({ auditLogs: [] })}
                >
                  Clear Logs
                </button>
              )}
            </div>
            <div className={styles.logContainer}>
              {this.state.auditLogs.length === 0 ? (
                <div className={styles.emptyLog}>No audit activities logged yet.</div>
              ) : (
                this.state.auditLogs.map((log, i) => (
                  <div key={`${log.timestamp.getTime()}-${i}`} className={`${styles.logEntry} ${styles[log.type]}`}>
                    <span className={styles.timestamp}>[{log.timestamp.toLocaleTimeString()}]</span>
                    <span className={styles.msg}>{log.message}</span>
                  </div>
                ))
              )}
            </div>
          </div>
        )}
      </section >
    );
  }
}
