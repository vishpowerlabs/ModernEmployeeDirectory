import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { IEmployee } from '../models/IEmployee';
import { GraphService } from '../services/GraphService';
import { KudosService } from '../services/KudosService';
import { Filters } from './Filters';
import { EmployeeGrid } from './EmployeeGrid';
import { EmployeeList } from './EmployeeList';
import { ProfileView } from './ProfileView';
import { AlphabetNav } from './AlphabetNav';
import { SkeletonCard } from './SkeletonCard';
import { ProfileEditForm } from './ProfileEditForm';
import { Icon } from '@fluentui/react/lib/Icon';
import { KudosPanel } from './KudosPanel';
import { MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

export interface IEmployeeDirectoryState {
  employees: IEmployee[];
  loading: boolean;
  view: 'home' | 'profile';
  selectedUser: IEmployee | null;
  search: string;
  filters: { [key: string]: string };
  presence: { [key: string]: string };
  nextLink?: string;
  isLoadingMore: boolean;
  letter: string;
  listMode: 'grid' | 'list';
  currentUser?: IEmployee;
  currentUserPhoto?: string;
  isPanelOpen: boolean;
  isCalloutVisible: boolean;
  message?: { text: string; type: 'success' | 'error' };
  isSaving?: boolean;
  kudosCounts: { [email: string]: number };
  isKudosPanelOpen: boolean;
  kudosPanelMode: 'give' | 'view';
  kudosTargetUser: IEmployee | null;
}

export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeDirectoryState> {
  private readonly _graphService: GraphService;
  private _kudosService: KudosService | undefined;
  // private readonly _photoRef = React.createRef<HTMLDivElement>(); // Removed unused ref

  constructor(props: IEmployeeDirectoryProps) {
    super(props);
    this.state = {
      employees: [],
      loading: true,
      view: 'home',
      selectedUser: null,
      search: '',
      filters: {},
      presence: {},
      letter: '',
      nextLink: undefined,
      isLoadingMore: false,
      listMode: 'grid',
      currentUser: undefined,
      currentUserPhoto: undefined,
      isPanelOpen: false,
      isCalloutVisible: false,
      kudosCounts: {},
      isKudosPanelOpen: false,
      kudosPanelMode: 'give',
      kudosTargetUser: null
    };
    this._graphService = new GraphService(this.props.context);
  }

  private readonly _getPresenceStatus = (userId: string): string | undefined => {
    return this.state.presence[userId];
  }

  public componentDidMount(): void {
    if (this.props.enableStarred && this.props.starredUsers && this.props.starredUsers.length > 0) {
      this.setState({ letter: '*' });
    }

    // Initialize Kudos Service
    if (this.props.enableKudos && this.props.kudosListId) {
      this._kudosService = new KudosService(this.props.context, this.props.kudosListId, {
        target: this.props.kudosTargetField,
        message: this.props.kudosMessageField,
        type: this.props.kudosTypeField,
        date: this.props.kudosDateField
      });
      this._fetchKudosCounts();
    }

    this._fetchEmployees();

    // Fetch current user details
    this._graphService.getCurrentUser().then(user => {
      if (user) {
        this.setState({ currentUser: user });
        this._graphService.getUserPhoto(user.id).then(photo => {
          if (photo) this.setState({ currentUserPhoto: photo });
        });
      }
    });
  }

  public componentDidUpdate(prevProps: IEmployeeDirectoryProps): void {
    if (
      prevProps.filterType !== this.props.filterType ||
      prevProps.filterValue !== this.props.filterValue ||
      prevProps.filterSecondaryValue !== this.props.filterSecondaryValue ||
      prevProps.excludeUsers !== this.props.excludeUsers
    ) {
      this._fetchEmployees();
    }

    if (
      prevProps.enableKudos !== this.props.enableKudos ||
      prevProps.kudosListId !== this.props.kudosListId ||
      prevProps.kudosTargetField !== this.props.kudosTargetField ||
      prevProps.kudosMessageField !== this.props.kudosMessageField ||
      prevProps.kudosTypeField !== this.props.kudosTypeField ||
      prevProps.kudosDateField !== this.props.kudosDateField
    ) {
      if (this.props.enableKudos && this.props.kudosListId) {
        this._kudosService = new KudosService(this.props.context, this.props.kudosListId, {
          target: this.props.kudosTargetField,
          message: this.props.kudosMessageField,
          type: this.props.kudosTypeField,
          date: this.props.kudosDateField
        });
        this._fetchKudosCounts();
      } else {
        this._kudosService = undefined;
        this.setState({ kudosCounts: {} });
      }
    }
  }

  private async _fetchKudosCounts(): Promise<void> {
    if (this._kudosService) {
      const counts = await this._kudosService.getKudosCounts();
      this.setState({ kudosCounts: counts });
    }
  }

  private async _fetchEmployees(): Promise<void> {
    try {
      const { filterType, filterValue, filterSecondaryValue } = this.props;
      const response = await this._graphService.getUsers(filterType, filterValue, filterSecondaryValue);
      // Filter out users without display name or basic info if needed
      const validUsers = response.users.filter(u => u.displayName);
      this.setState({ employees: validUsers, nextLink: response.nextLink, loading: false });
      this._fetchPresence(validUsers);
    } catch (error) {
      console.error("Error fetching employees", error);
      this.setState({ employees: [], loading: false });
    }
  }

  private async _fetchPresence(users: IEmployee[]): Promise<void> {
    const ids = users.map(u => u.id).filter(Boolean);
    if (ids.length === 0) return;

    try {
      const presenceList = await this._graphService.getPresence(ids);
      const presenceMap: { [key: string]: string } = {};
      presenceList.forEach(p => {
        presenceMap[p.id] = p.availability;
      });

      this.setState(prevState => ({
        presence: { ...prevState.presence, ...presenceMap }
      }));
    } catch (err) {
      console.warn("Presence fetch failed", err);
    }
  }

  private async _exportToCsv(): Promise<void> {
    const { employees } = this.state;
    if (employees.length === 0) return;

    const headers = ['DisplayName', 'JobTitle', 'Department', 'Email', 'Mobile', 'WorkPhone', 'Location', 'City', 'State', 'Country'];
    const csvContent = [
      headers.join(','),
      ...employees.map(u => [
        `"${u.displayName || ''}"`,
        `"${u.jobTitle || ''}"`,
        `"${u.department || ''}"`,
        `"${u.mail || ''}"`,
        `"${u.mobilePhone || ''}"`,
        `"${u.businessPhones ? u.businessPhones[0] : ''}"`,
        `"${u.officeLocation || ''}"`,
        `"${u.city || ''}"`,
        `"${u.state || ''}"`,
        `"${u.country || ''}"`
      ].join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'employee_directory_export.csv');
    document.body.appendChild(link);
    link.click();
    link.remove();
  }

  private async _loadMore(): Promise<void> {
    if (!this.state.nextLink) return;

    this.setState({ isLoadingMore: true });
    try {
      const response = await this._graphService.getUsers(
        this.props.filterType,
        this.props.filterValue,
        this.props.filterSecondaryValue,
        this.state.nextLink
      );

      const validUsers = response.users.filter(u => u.displayName);
      this.setState(prevState => ({
        employees: [...prevState.employees, ...validUsers],
        nextLink: response.nextLink,
        isLoadingMore: false
      }));
      this._fetchPresence(validUsers);
    } catch (error) {
      console.error("Error loading more users", error);
      this.setState({ isLoadingMore: false });
    }
  }

  // private readonly _handleUpdateProfile = (): void => { // Removed unused method
  //   this.setState({ isPanelOpen: true });
  // }

  private readonly _handleSaveProfile = async (data: Partial<IEmployee>): Promise<void> => {
    this.setState({ isSaving: true, message: undefined });
    try {
      await this._graphService.updateMyProfile(data);
      this.setState({
        isSaving: false,
        message: { text: 'Profile updated successfully!', type: 'success' }
      });

      // Refresh current user data
      const updatedUser = await this._graphService.getCurrentUser();
      if (updatedUser) {
        this.setState({ currentUser: updatedUser });
        const details = await this._graphService.getUserDetails(updatedUser.id);
        this.setState(prevState => ({
          currentUser: { ...prevState.currentUser, ...details } as IEmployee
        }));
      }

      // Close panel after a short delay to show success message
      setTimeout(() => {
        if (this.state.message?.type === 'success') {
          this._onPanelDismiss();
        }
      }, 2000);

    } catch (error) {
      console.error("Error saving profile", error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.setState({
        isSaving: false,
        message: { text: `Error updating profile: ${errorMessage}`, type: 'error' }
      });
    }
  }

  private readonly _onPanelDismiss = (): void => {
    this.setState({ isPanelOpen: false, message: undefined, isSaving: false });
  }

  private readonly _onGiveKudos = (user: IEmployee): void => {
    this.setState({
      isKudosPanelOpen: true,
      kudosPanelMode: 'give',
      kudosTargetUser: user
    });
  }

  private readonly _onViewKudos = (user: IEmployee): void => {
    this.setState({
      isKudosPanelOpen: true,
      kudosPanelMode: 'view',
      kudosTargetUser: user
    });
  }

  private readonly _onKudosDismiss = (): void => {
    this.setState({ isKudosPanelOpen: false });
  }

  private readonly _onKudosSubmitted = (): void => {
    this._fetchKudosCounts();
  }

  // Removed unused callout methods

  private _getUniqueValuesForField(field: string): string[] {
    const values = this.state.employees.map(e => (e as any)[field]).filter(Boolean);
    return Array.from(new Set(values as string[])).sort((a, b) => a.localeCompare(b));
  }

  private _getFilteredEmployees(): IEmployee[] {
    const { employees, search, filters, letter } = this.state;
    return employees.filter(e => {
      const searchLower = (search || '').toLowerCase();
      const matchSearch = !search ||
        ((e.displayName || '').toLowerCase().includes(searchLower)) ||
        ((e.jobTitle || '').toLowerCase().includes(searchLower)) ||
        ((e.mail || '').toLowerCase().includes(searchLower)) ||
        ((e.aboutMe || '').toLowerCase().includes(searchLower));

      const matchFilters = Object.keys(filters).every(field => {
        const filterValue = filters[field];
        return !filterValue || (e as any)[field] === filterValue;
      });

      const matchLetter = !letter ||
        (letter === '*' ? (this.props.starredUsers?.some(su => (su.email === e.mail || su.login === e.userPrincipalName))) :
          e.displayName?.toUpperCase().startsWith(letter.toUpperCase()));

      let matchExclude = true;
      if (this.props.excludeUsers) {
        const excludedTerms = (this.props.excludeUsers || '').split(',').map(t => t.trim().toLowerCase()).filter(Boolean);
        if (excludedTerms.length > 0) {
          const email = (e.mail || '').toLowerCase();
          const upn = (e.userPrincipalName || '').toLowerCase();
          const name = (e.displayName || '').toLowerCase();
          matchExclude = !excludedTerms.some(term => email.includes(term) || upn.includes(term) || name.includes(term));
        }
      }

      return matchSearch && matchFilters && matchLetter && matchExclude;
    });
  }


  private _renderHeader(): JSX.Element {
    const initials = this.state.currentUser?.displayName ? this.state.currentUser.displayName.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase() : '??';

    return (
      <div className={styles.header}>
        <div className={styles.container}>
          <div className={styles.headerInner}>
            <div className={styles.headerContent}>
              <h1 style={{ fontSize: `${this.props.titleFontSize || 28}px` }}>{this.props.title || 'Employee Directory'}</h1>
              <p style={{ fontSize: `${this.props.descriptionFontSize || 14}px` }}>{this.props.description || 'Find and connect with colleagues'}</p>
            </div>

            {/* Current User Trigger */}
            {this.state.currentUser && (
              <button
                className={styles.headerUser}
                onClick={() => this.setState({ isPanelOpen: true })}
                title="My Profile"
                style={{
                  border: 'none', padding: 0, font: 'inherit', cursor: 'pointer',
                  width: '56px', height: '56px', borderRadius: '50%', background: 'var(--m-bg-white)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '24px'
                }}
              >
                {this.state.currentUserPhoto ? (
                  <img src={this.state.currentUserPhoto} alt={this.state.currentUser.displayName} className={styles.userPhoto} />
                ) : (
                  <span className={styles.userInitial}>{initials}</span>
                )}
                <div className={styles.userOnlineStatus}></div>
              </button>
            )}
          </div>
        </div>

        {/* Custom Side Panel (My Profile) */}
        <button
          className={`${styles.overlay} ${this.state.isPanelOpen ? styles.open : ''}`}
          onClick={this._onPanelDismiss}
          style={{
            border: 'none', padding: 0, margin: 0, cursor: 'default',
            position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh',
            background: 'rgba(0,0,0,0.3)', zIndex: 999,
            display: this.state.isPanelOpen ? 'block' : 'none'
          }}
          aria-label="Close profile panel"
        ></button>
        <div className={`${styles.sidePanel} ${this.state.isPanelOpen ? styles.open : ''}`}>
          <div className={styles.sidePanelHeader}>
            <h2>My Profile</h2>
            <button className={styles.closePanel} onClick={this._onPanelDismiss}>✕</button>
          </div>
          <div className={styles.sidePanelContent}>
            {this.state.currentUser && (
              <ProfileEditForm
                employee={this.state.currentUser}
                graphService={this._graphService}
                editFields={this.props.editFields}
                onCancel={this._onPanelDismiss}
                onSave={this._handleSaveProfile}
              />
            )}
            {this.state.message && (
              <MessageBar
                messageBarType={this.state.message.type === 'success' ? MessageBarType.success : MessageBarType.error}
                onDismiss={() => this.setState({ message: undefined })}
                style={{ marginTop: '15px' }}
              >
                {this.state.message.text}
              </MessageBar>
            )}
            {this.state.isSaving && (
              <div style={{ marginTop: '15px' }}>
                <Spinner size={SpinnerSize.medium} label="Saving changes..." labelPosition="right" />
              </div>
            )}
            <div style={{ textAlign: 'center', marginTop: '30px', color: '#999', fontSize: '10px' }}>v0.0.19</div>
          </div>
        </div>
      </div>
    );
  }

  private _renderHomeView(filteredEmployees: IEmployee[], totalEmployees: number, sectionTitle: string, filterOptions: { [key: string]: string[] }): JSX.Element {
    return (
      <div className={`${styles.viewContainer} ${styles.active}`}>
        <div className={styles.searchSection}>
          <div className={styles.container}>
            <Filters
              onSearch={(s) => this.setState({ search: s })}
              onFilter={(f) => this.setState({ filters: f })}
              filterFields={this.props.filterFields}
              filterOptions={filterOptions}
              searchQuery={this.state.search}
            />
          </div>
        </div>

        <div className={styles.alphabetNav}>
          <div className={styles.container}>
            <AlphabetNav
              selectedLetter={this.state.letter}
              onLetterClick={(l) => this.setState({ letter: l })}
              enableStarred={this.props.enableStarred}
            />
          </div>
        </div>

        <div className={styles.resultsHeader}>
          <div className={styles.container}>
            <div className={styles.resultsInner}>
              <div className={styles.resultsInfo}>
                <h2>{sectionTitle}</h2>
                <p>Found {filteredEmployees.length} employees</p>
              </div>
              <div className={styles.actionButtons}>
                <button
                  className={`${styles.actionBtn} ${this.state.listMode === 'grid' ? styles.active : ''}`}
                  onClick={() => this.setState({ listMode: 'grid' })}
                  title="Grid View"
                >
                  <Icon iconName="GridViewSmall" />
                  Grid
                </button>
                <button
                  className={`${styles.actionBtn} ${this.state.listMode === 'list' ? styles.active : ''}`}
                  onClick={() => this.setState({ listMode: 'list' })}
                  title="List View"
                >
                  <Icon iconName="CollapseMenu" />
                  List
                </button>
                <button className={styles.actionBtn} onClick={() => this._exportToCsv()} title="Export CSV">
                  <Icon iconName="Download" />
                  Export
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className={`${styles.mainContent} ${styles.container}`} style={{ marginTop: '24px' }}>
          {this.state.listMode === 'list' ? (
            <EmployeeList
              employees={filteredEmployees}
              onUserSelect={(u) => this.setState({ selectedUser: u, view: 'profile' })}
              graphService={this._graphService}
              presence={this.state.presence}
              kudosCounts={this.state.kudosCounts}
              onGiveKudos={this._onGiveKudos}
              onViewKudos={this._onViewKudos}
              isKudosEnabled={!!(this.props.enableKudos && this.props.kudosListId)}
              visualTheme={this.props.visualTheme}
            />
          ) : (
            <EmployeeGrid
              employees={filteredEmployees}
              onUserSelect={(u) => this.setState({ selectedUser: u, view: 'profile' })}
              graphService={this._graphService}
              presence={this.state.presence}
              kudosCounts={this.state.kudosCounts}
              onGiveKudos={this._onGiveKudos}
              onViewKudos={this._onViewKudos}
              isKudosEnabled={!!(this.props.enableKudos && this.props.kudosListId)}
              visualTheme={this.props.visualTheme}
            />
          )}

          {this.state.nextLink && !this.state.loading && (
            <div style={{ textAlign: 'center', marginTop: '40px', paddingBottom: '40px' }}>
              <button
                className={styles.profileBtn}
                onClick={() => this._loadMore()}
                disabled={this.state.isLoadingMore}
                style={{ minWidth: '200px', height: '44px' }}
              >
                {this.state.isLoadingMore ? 'Loading...' : 'Load More Employees'}
              </button>
            </div>
          )}
        </div>
      </div>
    );
  }

  private _getThemeClass(): string {
    switch (this.props.themeTemplate) {
      case 'default': return '';
      case 'berry': return styles.themeBerry;
      case 'orange': return styles.themeOrange;
      case 'grape': return styles.themeGrape;
      case 'forest': return styles.themeForest;
      case 'ocean': return styles.themeOcean;
      case 'midnight': return styles.themeMidnight;
      case 'mint': return styles.themeMint;
      default: return '';
    }
  }

  public render(): React.ReactElement<IEmployeeDirectoryProps> {
    const { employees, loading, view, selectedUser, letter } = this.state;
    const filterOptions: { [key: string]: string[] } = {};

    const filteredEmployees = this._getFilteredEmployees();

    if (this.props.filterFields) {
      this.props.filterFields.forEach(field => {
        filterOptions[field] = this._getUniqueValuesForField(field);
      });
    }

    const totalEmployees = employees.length;
    let sectionTitle = 'All Employees';
    if (letter === '*') {
      sectionTitle = 'Key People';
    } else if (letter) {
      sectionTitle = `Employees - ${letter}`;
    }

    const themeClass = this._getThemeClass();

    if (loading) {
      return (
        <div className={styles.mainWrapper}>
          <div className={styles.header}>
            <div className={styles.container}>
              <div className={styles.headerInner}>
                <div className={styles.headerContent}>
                  <h1>{this.props.title || 'Employee Directory'}</h1>
                  <p>{this.props.description || 'Connect with colleagues across the organization'}</p>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.employeeSection}>
            <div className={styles.employeeGrid}>
              {[1, 2, 3, 4, 5, 6].map(i => <SkeletonCard key={i} />)}
            </div>
          </div>
        </div>
      );
    }

    let visualThemeClass = '';
    switch (this.props.visualTheme) {
      case 'modern-scrolling-premium': visualThemeClass = styles.themeScrolling; break;
      case 'modern-tabbed-premium': visualThemeClass = styles.themeTabbed; break;
    }

    return (
      <div
        className={`${styles.employeeDirectory} ${themeClass} ${visualThemeClass}`}
        style={{
          '--subHeadingFontSize': `${this.props.subHeadingFontSize || 20}px`,
          '--contentFontSize': `${this.props.contentFontSize || 14}px`,
          '--headerPadding': `${this.props.headerPadding || 40}px`,
          '--profileHeaderPadding': `${this.props.profileHeaderPadding || 30}px`
        } as React.CSSProperties}
      >
        <div className={styles.mainWrapper}>
          {this._renderHeader()}

          {view === 'home' && this._renderHomeView(filteredEmployees, totalEmployees, sectionTitle, filterOptions)}

          {/* Profile View */}
          {view === 'profile' && selectedUser && (
            <ProfileView
              employee={selectedUser}
              onUserSelect={(u) => this.setState({ selectedUser: u })}
              graphService={this._graphService}
              visualTheme={this.props.visualTheme}
              enableKudos={this.props.enableKudos}
              title={this.props.title}
              kudosCounts={this.state.kudosCounts}
              onGiveKudos={this._onGiveKudos}
              onViewKudos={this._onViewKudos}
              presenceStatus={this._getPresenceStatus(selectedUser.id)}
              onBackToHome={() => this.setState({ view: 'home', selectedUser: null })}
            />
          )}
        </div>

        {/* Kudos Panel */}
        <KudosPanel
          isOpen={this.state.isKudosPanelOpen}
          onDismiss={this._onKudosDismiss}
          mode={this.state.kudosPanelMode}
          targetUser={this.state.kudosTargetUser}
          currentUser={this.state.currentUser}
          kudosService={this._kudosService}
          onKudosSubmitted={this._onKudosSubmitted}
          themeClass={themeClass}
        />
      </div>
    );
  }
}
