import * as React from 'react';
import styles from './ModernEmployeeDirectory.module.scss';
import type { IModernEmployeeDirectoryProps } from './IModernEmployeeDirectoryProps';
import { TopBar } from './Shared/TopBar';
import { SidePanel } from './Shared/SidePanel';
import { DirectoryHome, IEmployee } from './Home/DirectoryHome';
import { ProfileScrolling } from './Profile/ProfileScrolling';
import { ProfileTabbed } from './Profile/ProfileTabbed';
import { MockDataService } from '../services/MockDataService';
import { KudosService } from '../services/KudosService';
import { GraphService, IGraphUser } from '../services/GraphService';
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
  nextPageLink: string | null;
  prevPageLinks: string[];
  lastFetchedLink: string | null;
  selectedLetter: string | null;
  dynamicFilterData: { [key: string]: string[] };
}

const MOCK_CURRENT_USER = {
  initials: 'CU',
  name: 'Current User',
  email: 'user@modern.com'
};

export default class ModernEmployeeDirectory extends React.Component<IModernEmployeeDirectoryProps, IModernEmployeeDirectoryState> {
  private kudosService: KudosService | null = null;
  private readonly graphService: GraphService | null = null;
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
      nextPageLink: null,
      prevPageLinks: [],
      lastFetchedLink: null,
      selectedLetter: 'STAR',
      dynamicFilterData: {}
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
        console.error(`[ModernEmployeeDirectory] Error fetching unique values for ${field}:`, error);
        dynamicFilterData[field] = [];
      }
    });

    await Promise.all(fetchPromises);
    this.setState({ dynamicFilterData });
  }

  public componentDidMount(): void {
    void (async (): Promise<void> => {
      this._injectWorkbenchStyles();

      try {
        const currentUser = this.props.context.pageContext.user;
        console.log('[ModernEmployeeDirectory] Current user from context:', currentUser);

        this.setState({
          currentUser: {
            id: currentUser.loginName || currentUser.email,
            initials: this._getInitials(currentUser.displayName),
            name: currentUser.displayName,
            email: currentUser.loginName || currentUser.email
          }
        }, async () => {
          if (this.graphService) {
            try {
              // Fetch actual Graph UUID for 'me' to ensure matching in Org Chart
              const me = await this.graphService.getCurrentUser();
              if (me) {
                this.setState(prevState => ({
                  currentUser: {
                    ...prevState.currentUser,
                    id: me.id,
                    jobTitle: me.jobTitle,
                    mobilePhone: me.mobilePhone
                  }
                }));
                console.log('[ModernEmployeeDirectory] ✅ Synced current user Graph ID:', me.id);
              }

              const photoUrl = await this.graphService.getUserPhoto('me');
              if (photoUrl) {
                this.setState(prevState => ({
                  currentUser: { ...prevState.currentUser, photoUrl }
                }));
                console.log('[ModernEmployeeDirectory] ✅ Fetched current user photo');
              }
            } catch (error) {
              console.warn('[ModernEmployeeDirectory] No photo for current user:', error);
            }
          }
        });
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error getting current user:', error);
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
    })();
  }

  public componentDidUpdate(prevProps: IModernEmployeeDirectoryProps): void {
    void (async (): Promise<void> => {
      if (prevProps.dataSource !== this.props.dataSource) {
        console.log('[ModernEmployeeDirectory] Data source changed, reloading employees...');
        await this._loadEmployees();
      }

      if (JSON.stringify(prevProps.homePageFilterFields) !== JSON.stringify(this.props.homePageFilterFields)) {
        console.log('[ModernEmployeeDirectory] Filter fields changed, fetching new values...');
        await this._fetchDynamicFilterValues();
      }

      if (prevProps.profileLayout !== this.props.profileLayout && this.state.selectedEmployee) {
        console.log('[ModernEmployeeDirectory] Layout changed, switching profile view...');
        this.setState({
          currentView: this.props.profileLayout === 'scroll' ? 'PROFILE_SCROLL' : 'PROFILE_TAB'
        });
      }

      if (prevProps.enableKudos !== this.props.enableKudos ||
        prevProps.kudosListId !== this.props.kudosListId ||
        prevProps.kudosRecipientColumn !== this.props.kudosRecipientColumn ||
        prevProps.kudosAuthorColumn !== this.props.kudosAuthorColumn ||
        prevProps.kudosMessageColumn !== this.props.kudosMessageColumn ||
        prevProps.kudosBadgeTypeColumn !== this.props.kudosBadgeTypeColumn) {

        console.log('[ModernEmployeeDirectory] Kudos settings changed, re-initializing service...');
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

      if (prevProps.filterType !== this.props.filterType ||
        prevProps.filterValue !== this.props.filterValue ||
        prevProps.filterSecondaryValue !== this.props.filterSecondaryValue) {
        console.log('[ModernEmployeeDirectory] Filter settings changed, reloading employees...');
        await this._loadEmployees();
      }
    })();
  }

  private async _loadEmployees(): Promise<void> {
    const { dataSource } = this.props;
    const token = ++this._loadToken;
    this.setState({ employeesLoading: true });

    try {
      if (dataSource === 'graph' && this.graphService) {
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

        void this._fetchEmployeePhotos(employees, token);
        void this._fetchManagerInfo(employees, token);
      } else {
        const employees = MockDataService.getEmployees();
        this.setState({ employees, employeesLoading: false });
      }
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error loading employees:', error);
      const employees = MockDataService.getEmployees();
      this.setState({ employees, employeesLoading: false });
    }
  }

  private readonly _handleLoadMore = async (): Promise<void> => {
    const { nextPageLink } = this.state;
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

      void this._fetchEmployeePhotos(newEmployees, token);
      void this._fetchManagerInfo(newEmployees, token);
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error loading more employees:', error);
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

      void this._fetchEmployeePhotos(newEmployees, token);
      void this._fetchManagerInfo(newEmployees, token);
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error loading next page:', error);
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

      void this._fetchEmployeePhotos(employees, token);
      void this._fetchManagerInfo(employees, token);
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error loading previous page:', error);
      this.setState({ employeesLoading: false });
    }
  }

  private readonly _handleLetterChange = async (letter: string | null): Promise<void> => {
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
          console.warn('[ModernEmployeeDirectory] No photo for:', emp.displayName, error);
        }
      }));
    }
  }

  private async _fetchManagerInfo(employees: IEmployee[], token: number): Promise<void> {
    if (!this.graphService) return;

    const batchSize = 10;
    for (let i = 0; i < employees.length; i += batchSize) {
      if (token !== this._loadToken) return;
      const batch = employees.slice(i, i + batchSize);

      await Promise.all(batch.map(async (emp) => {
        try {
          const manager = await this.graphService!.getUserManager(emp.id);
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
                  isOnline: false,
                  _isOrgChartNode: true
                };
                updatedEmployees = [...updatedEmployees, managerEmp];
              }

              return { employees: updatedEmployees };
            });
          }
        } catch (error) {
          console.warn('[ModernEmployeeDirectory] No manager for:', emp.displayName, error);
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
      isOnline: false
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

    console.log('[ModernEmployeeDirectory] Loading Hall of Fame (Star) employees...');
    let hofEmails: string[] = [];
    if (this.kudosService && this.props.minKudosCount > 0) {
      try {
        hofEmails = await this.kudosService.getHallOfFameCandidates(this.props.minKudosCount);
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error getting HOF candidates:', error);
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
      console.error('[ModernEmployeeDirectory] Error loading HOF details:', error);
      return [];
    }
  }

  private _injectWorkbenchStyles(): void {
    if (document.getElementById('dev-workbench-fullwidth')) return;

    const style = document.createElement('style');
    style.id = 'dev-workbench-fullwidth';
    style.textContent = `
      .p_ZwSiC_hHQBj, .CanvasComponent, .LCS, .CanvasZone,
      div[data-automation-id="CanvasZone"], .CanvasSection {
        max-width: 100% !important;
        width: 100% !important;
      }
      .CanvasComponent.LCS .CanvasZone {
        max-width: 100% !important;
        width: 100% !important;
        padding: 0 !important;
        margin: 0 auto !important;
      }
    `;
    document.head.appendChild(style);
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
        console.error('[ModernEmployeeDirectory] Error fetching detailed profile:', error);
      }
    }

    let kudosCount = 0;
    if (this.props.enableKudos && this.kudosService && employee.mail) {
      try {
        kudosCount = await this.kudosService.getKudosCount(employee.mail);
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching kudos count:', error);
      }
    }

    this.setState({
      selectedEmployee: detailedEmployee,
      currentView: this.props.profileLayout === 'scroll' ? 'PROFILE_SCROLL' : 'PROFILE_TAB',
      selectedEmployeeKudosCount: kudosCount
    });
  }

  private readonly _handleKudosForEmployee = async (employee: IEmployee): Promise<void> => {
    let kudosCount = 0;
    if (this.props.enableKudos && this.kudosService && employee.mail) {
      try {
        kudosCount = await this.kudosService.getKudosCount(employee.mail);
      } catch (error) {
        console.error('[ModernEmployeeDirectory] Error fetching kudos count:', error);
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

    if (mode === 'updateProfile') {
      void this._loadFreshProfileData();
    }

    if (mode === 'kudos' && this.props.enableKudos) {
      const userId = this.state.selectedEmployee?.mail || this.state.currentUser.email;
      void this._loadKudos(userId);
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
      console.error('[ModernEmployeeDirectory] Error fetching fresh profile data:', error);
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
      console.error('[ModernEmployeeDirectory] Error loading kudos:', error);
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
        await this._loadKudos(recipient.mail);
      }
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error giving kudos:', error);
    }
  }

  private readonly _handleUpdateProfile = async (updatedFields: any): Promise<boolean> => {
    if (!this.graphService) return false;

    try {
      const result = await this.graphService.updateUserProfile(updatedFields);
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
      }
      return result.success;
    } catch (error) {
      console.error('[ModernEmployeeDirectory] Error updating profile:', error);
      return false;
    }
  }

  public render(): React.ReactElement<IModernEmployeeDirectoryProps> {
    const { hasTeamsContext, mainHeadingSize, subHeadingSize, contentHeadingSize } = this.props;
    const { currentView, selectedEmployee, isSidePanelOpen } = this.state;

    const fontSizeStyles = {
      '--main-heading-size': `${mainHeadingSize}px`,
      '--sub-heading-size': `${subHeadingSize}px`,
      '--content-heading-size': `${contentHeadingSize}px`
    } as React.CSSProperties;

    return (
      <section className={`${styles.modernEmployeeDirectory} ${hasTeamsContext ? styles.teams : ''}`} style={fontSizeStyles} >
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
                this._handleSelectEmployee(myEmployee);
              } else {
                this._toggleSidePanel(mode);
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
            />
          )}
        </div>
      </section >
    );
  }
}
