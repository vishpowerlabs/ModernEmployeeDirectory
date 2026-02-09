import * as React from 'react';
import styles from './DirectoryHome.module.scss';


export interface IDirectoryHomeProps {
    employees: IEmployee[];
    onSelectEmployee: (employee: IEmployee) => void;
    onKudosForEmployee: (employee: IEmployee) => void;
    enableKudos?: boolean;
    onLoadMore?: () => void;
    hasMore?: boolean;
    loading?: boolean;
    paginationType?: 'loadMore' | 'prevNext';
    onNextPage?: () => void;
    onPrevPage?: () => void;
    hasPrev?: boolean;
    selectedLetter?: string | null;
    onLetterChange?: (letter: string | null) => void;
    homePageFilterFields?: string[];
    dynamicFilterData?: { [key: string]: string[] };
}

export interface IEmployee {
    // Core Microsoft Graph User properties
    id: string;                          // Graph: id (unique identifier)
    displayName: string;                 // Graph: displayName
    givenName?: string;                  // Graph: givenName (first name)
    surname?: string;                    // Graph: surname (last name)
    jobTitle?: string;                   // Graph: jobTitle
    department?: string;                 // Graph: department
    officeLocation?: string;             // Graph: officeLocation
    city?: string;                       // Graph: city
    state?: string;                      // Graph: state
    country?: string;                    // Graph: country
    mail?: string;                       // Graph: mail (primary email)
    userPrincipalName?: string;          // Graph: userPrincipalName (UPN)
    mobilePhone?: string;                // Graph: mobilePhone
    businessPhones?: string[];           // Graph: businessPhones (array)

    // Manager relationship
    managerId?: string;                  // Custom: manager's id (from Graph manager endpoint)

    // Colleagues/Peers (People I Work With)
    colleagues?: string[];               // Custom: array of colleague user IDs

    // Extended profile fields from Graph
    aboutMe?: string;                    // Graph: aboutMe (user bio/description)
    skills?: string[];                   // Graph: skills (array of skills)
    interests?: string[];                // Graph: interests
    responsibilities?: string[];         // Graph: responsibilities
    projects?: string[];                 // Graph: projects (current projects)
    pastProjects?: string[];             // Graph: pastProjects
    photoUrl?: string;                   // Graph: photo URL (blob URL for profile photo)

    // Presence (from Graph Presence API)
    presence?: {
        availability?: string;           // Graph Presence: available, busy, away, offline, etc.
        activity?: string;               // Graph Presence: Available, InACall, InAMeeting, etc.
    };

    // Computed/Helper properties (not from Graph, computed client-side)
    initials?: string;                   // Computed from displayName or givenName/surname
    isOnline?: boolean;                  // Computed from presence.availability
    isFeatured?: boolean;                // Flag for Hall of Fame manually featured
    isTopKudos?: boolean;                // Flag for Hall of Fame kudos threshold
    _isOrgChartNode?: boolean;           // Internal flag for nodes injected specifically for OrgChart
}

const EmployeeCard: React.FC<{
    emp: IEmployee;
    viewMode: 'grid' | 'list';
    onSelectEmployee: (emp: IEmployee) => void;
    onKudosForEmployee: (emp: IEmployee) => void;
    enableKudos?: boolean;
}> = ({ emp, viewMode, onSelectEmployee, onKudosForEmployee, enableKudos }) => {
    return (
        <div className={`${styles.employeeCard} ${viewMode === 'list' ? styles.listView : ''}`}>
            <div className={styles.badgeContainer}>
                {emp.isFeatured && (
                    <span className={styles.vipBadge} title="Featured VIP">
                        <span className={styles.badgeIcon}>💎</span> VIP
                    </span>
                )}
                {emp.isTopKudos && (
                    <span className={styles.kudosBadge} title="Kudos Star recipient">
                        <span className={styles.badgeIcon}>⭐</span> Star
                    </span>
                )}
            </div>
            <div className={styles.cardHeader}>
                <button
                    className={styles.avatar}
                    onClick={() => onSelectEmployee(emp)}
                    title="View Profile"
                >
                    {emp.photoUrl ? (
                        <img src={emp.photoUrl} alt={emp.displayName} style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover' }} />
                    ) : (
                        emp.initials
                    )}
                    {emp.isOnline && <div className={styles.onlineStatus}></div>}
                </button>
                <div className={styles.employeeInfo}>
                    <h3 className={styles.employeeName}>
                        <button onClick={() => onSelectEmployee(emp)} title="View Profile">
                            {emp.displayName}
                        </button>
                    </h3>
                    <div className={styles.employeeTitle}>{emp.jobTitle || 'N/A'}</div>
                    <div className={styles.employeeDept}>{emp.department || 'N/A'}</div>
                </div>
            </div>

            <div className={styles.cardDetails}>
                <div className={styles.detailRow}>
                    <span>📍</span> {emp.city && emp.state ? `${emp.city}, ${emp.state}` : emp.officeLocation || 'N/A'}
                </div>
                <div className={styles.detailRow}>
                    <span>✉️</span> {emp.mail || emp.userPrincipalName || 'N/A'}
                </div>
                <div className={styles.detailRow}>
                    <span>📞</span> {emp.businessPhones?.[0] || emp.mobilePhone || 'N/A'}
                </div>
            </div>

            <div className={styles.cardActions}>
                <button className={styles.iconBtn} title="Email">✉️</button>
                <button className={styles.iconBtn} title="Chat">💬</button>
                <button className={styles.iconBtn} title="Call">📞</button>
                {enableKudos && (
                    <button
                        className={styles.iconBtn}
                        title="Give Kudos"
                        onClick={(e) => {
                            e.stopPropagation();
                            onKudosForEmployee(emp);
                        }}
                    >
                        ⭐
                    </button>
                )}
            </div>
        </div>
    );
};

const AlphabetNav: React.FC<{
    selectedLetter: string | null | undefined;
    onLetterChange?: (letter: string | null) => void;
}> = ({ selectedLetter, onLetterChange }) => (
    <div className={styles.alphabetNav}>
        <div className="container">
            <h3>Filter by Name</h3>
            <div className={styles.letters}>
                <button
                    className={`${styles.letter} ${styles.star} ${selectedLetter === 'STAR' ? styles.active : ''}`}
                    onClick={() => onLetterChange?.('STAR')}
                >
                    ⭐
                </button>
                <button
                    className={`${styles.letter} ${selectedLetter === null ? styles.active : ''}`}
                    onClick={() => onLetterChange?.(null)}
                >
                    All
                </button>
                {'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('').map(letter => (
                    <button
                        key={letter}
                        className={`${styles.letter} ${selectedLetter === letter ? styles.active : ''}`}
                        onClick={() => onLetterChange?.(letter)}
                    >
                        {letter}
                    </button>
                ))}
            </div>
        </div>
    </div>
);

export const DirectoryHome: React.FunctionComponent<IDirectoryHomeProps> = (props) => {
    const {
        employees, onSelectEmployee, onKudosForEmployee, enableKudos,
        onLoadMore, hasMore, loading,
        paginationType, onNextPage, onPrevPage, hasPrev,
        selectedLetter, onLetterChange,
        homePageFilterFields, dynamicFilterData
    } = props;
    const [viewMode, setViewMode] = React.useState<'grid' | 'list'>('grid');
    const [searchTerm, setSearchTerm] = React.useState('');
    const [activeFilters, setActiveFilters] = React.useState<{ [key: string]: string }>({});

    const _handleFilterChange = (field: string, value: string): void => {
        setActiveFilters(prev => ({
            ...prev,
            [field]: value === 'ALL' ? '' : value
        }));
    };

    const filteredEmployees = employees.filter(emp => {
        if (emp._isOrgChartNode) return false;

        // Search text matching (Name, Job Title, Department)
        const searchLower = searchTerm.toLowerCase();
        const matchesSearch = !searchTerm ||
            emp.displayName.toLowerCase().includes(searchLower) ||
            (emp.jobTitle?.toLowerCase() || '').includes(searchLower) ||
            (emp.department?.toLowerCase() || '').includes(searchLower);

        if (!matchesSearch) return false;

        // Dynamic multi-field filter matching
        return Object.keys(activeFilters).every(field => {
            const filterValue = activeFilters[field];
            if (!filterValue) return true; // Filter is cleared

            const empValue = (emp as any)[field];
            return empValue === filterValue;
        });
    });

    return (
        <>
            {/* Search Section */}
            <div className={styles.searchSection}>
                <div className="container">
                    <div className={styles.searchBar}>
                        <input
                            type="text"
                            className={styles.searchInput}
                            placeholder="Search employees by name, title, or department..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                        />
                        {homePageFilterFields?.map(field => (
                            <select
                                key={field}
                                className={styles.filterDropdown}
                                value={activeFilters[field] || 'ALL'}
                                onChange={(e) => _handleFilterChange(field, e.target.value)}
                            >
                                <option value="ALL">All {field.charAt(0).toUpperCase() + field.slice(1).replace(/([A-Z])/g, ' $1').trim()}</option>
                                {dynamicFilterData?.[field]?.map(val => (
                                    <option key={val} value={val}>{val}</option>
                                ))}
                            </select>
                        ))}
                        <div className={styles.viewToggle}>
                            <button
                                className={`${styles.toggleBtn} ${viewMode === 'grid' ? styles.active : ''}`}
                                onClick={() => setViewMode('grid')}
                                title="Grid View"
                            >
                                ⊞
                            </button>
                            <button
                                className={`${styles.toggleBtn} ${viewMode === 'list' ? styles.active : ''}`}
                                onClick={() => setViewMode('list')}
                                title="List View"
                            >
                                ☰
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            {/* Alphabet Navigation */}
            <AlphabetNav
                selectedLetter={selectedLetter}
                onLetterChange={onLetterChange}
            />

            {/* Employee Grid/List */}
            <div className={styles.employeeSection}>
                <div className="container">
                    {filteredEmployees.length === 0 && !loading && (
                        <div className={styles.noResults}>
                            <div className={styles.noResultsIcon}>🔍</div>
                            <h3>No employees found</h3>
                            <p>We couldn't find any employees matching your search or filters. Try adjusting your criteria or clearing the search.</p>
                        </div>
                    )}

                    <div className={`${styles.employeeGrid} ${viewMode === 'list' ? styles.listView : ''}`}>
                        {filteredEmployees.map(emp => (
                            <EmployeeCard
                                key={emp.id}
                                emp={emp}
                                viewMode={viewMode}
                                onSelectEmployee={onSelectEmployee}
                                onKudosForEmployee={onKudosForEmployee}
                                enableKudos={enableKudos}
                            />
                        ))}
                    </div>

                    {/* Pagination Section */}
                    {((hasMore || hasPrev) && filteredEmployees.length > 0) && (
                        <div className={styles.paginationSection}>
                            {paginationType === 'prevNext' ? (
                                <div className={styles.prevNextContainer}>
                                    {hasPrev && (
                                        <button
                                            className={styles.prevNextBtn}
                                            onClick={onPrevPage}
                                            disabled={loading}
                                        >
                                            Previous
                                        </button>
                                    )}
                                    {hasMore && (
                                        <button
                                            className={styles.prevNextBtn}
                                            onClick={onNextPage}
                                            disabled={loading}
                                        >
                                            {loading ? 'Loading...' : 'Next'}
                                        </button>
                                    )}
                                </div>
                            ) : (
                                hasMore && (
                                    <button
                                        className={styles.loadMoreBtn}
                                        onClick={onLoadMore}
                                        disabled={loading}
                                    >
                                        {loading ? 'Loading...' : 'Load More Employees'}
                                    </button>
                                )
                            )}
                        </div>
                    )}
                </div>
            </div>
        </>
    );
};
