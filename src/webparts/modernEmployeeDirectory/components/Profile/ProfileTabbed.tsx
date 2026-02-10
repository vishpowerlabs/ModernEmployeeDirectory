import * as React from 'react';
import styles from './ProfileTabbed.module.scss';
import { IEmployee } from '../Home/DirectoryHome';
import { ContactItem, ActionButton, ProfileSection, ProfileTopBar, ProfileOrgChart, handleProfileContactClick } from './ProfileCommon';

export interface IProfileTabbedProps {
    employee: IEmployee;
    employees: IEmployee[];  // All employees for org chart
    onBack: () => void;
    onKudosClick: () => void;
    kudosCount?: number;
    onEmployeeSelect?: (employee: IEmployee) => void;  // Callback when clicking employee in org chart
    orgChartLayout?: 'vertical' | 'horizontal' | 'compact';
    onAuditLog?: (activity: string, target: string, details: any) => void;
}

export const ProfileTabbed: React.FunctionComponent<IProfileTabbedProps> = (props) => {
    const { employee, employees, onBack, kudosCount = 0, onEmployeeSelect, orgChartLayout, onAuditLog } = props;
    const [activeTab, setActiveTab] = React.useState('overview');

    const _handleTabChange = (tab: string): void => {
        if (activeTab === tab) return;
        setActiveTab(tab);
    };

    const _handleContactClick = (type: 'Email' | 'Teams' | 'Call'): void => {
        handleProfileContactClick(type, employee, onAuditLog);
    };

    return (
        <div className={styles.profileTabbed}>
            <ProfileTopBar onBack={onBack} styles={styles} />

            <div className={styles.container}>
                {/* Profile Header */}
                <div className={styles.profileHeader}>
                    <div className={styles.profileLeft}>
                        <div className={styles.avatarLarge}>
                            {employee.photoUrl ? (
                                <img src={employee.photoUrl} alt={employee.displayName} style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover' }} />
                            ) : (
                                employee.initials
                            )}
                            {employee.isOnline && <div className={styles.onlineStatus}></div>}
                        </div>
                        <div className={styles.profileInfo}>
                            <h1>{employee.displayName}</h1>
                            <p className={styles.profileTitle}>{employee.jobTitle || 'N/A'}</p>
                            <div className={styles.profileLocation}>
                                <ContactItem icon="✉️" text={employee.mail || employee.userPrincipalName || 'N/A'} className={styles.headerContactItem} />
                                <ContactItem icon="📞" text={employee.businessPhones?.[0] || employee.mobilePhone || 'N/A'} className={styles.headerContactItem} />
                                <ContactItem icon="📍" text={employee.city && employee.state ? `${employee.city}, ${employee.state}` : employee.officeLocation || 'N/A'} className={styles.headerContactItem} />
                            </div>
                        </div>
                    </div>
                    <div className={styles.quickActions}>
                        <ActionButton icon="✉️" label="Email" onClick={() => _handleContactClick('Email')} className={`${styles.actionBtn} ${styles.primary}`} />
                        <ActionButton icon="💬" label="Teams" onClick={() => _handleContactClick('Teams')} className={styles.actionBtn} />
                        <ActionButton icon="📞" label="Call" onClick={() => _handleContactClick('Call')} className={styles.actionBtn} />
                        <ActionButton
                            icon="⭐"
                            label="Give Kudos"
                            onClick={() => {
                                if (onAuditLog) onAuditLog('Kudos Interaction', employee.displayName || employee.mail || '', { source: 'ProfileButton', action: 'OpenKudosPanel' });
                                if (props.onKudosClick) props.onKudosClick();
                            }}
                            className={styles.btnKudos}
                            variant="kudos"
                            kudosCount={kudosCount}
                            kudosCountClassName={styles.kudosCount}
                        />
                    </div>
                </div>

                {/* Tab Navigation */}
                <div className={styles.tabNavigation}>
                    <button
                        className={`${styles.tabBtn} ${activeTab === 'overview' ? styles.active : ''}`}
                        onClick={() => _handleTabChange('overview')}
                    >
                        📋 Overview
                    </button>

                    <button
                        className={`${styles.tabBtn} ${activeTab === 'details' ? styles.active : ''}`}
                        onClick={() => _handleTabChange('details')}
                    >
                        📝 More Details
                    </button>
                    <button
                        className={`${styles.tabBtn} ${activeTab === 'organization' ? styles.active : ''}`}
                        onClick={() => _handleTabChange('organization')}
                    >
                        🏢 Organization
                    </button>
                </div>

                {/* Tab Content */}
                {activeTab === 'overview' && (
                    <div className={`${styles.tabContent} ${styles.active}`}>
                        <ProfileSection
                            icon="👤"
                            title="Professional Bio"
                            cardClassName={styles.contentCard}
                            headerClassName={styles.cardHeader}
                            iconClassName={styles.cardIcon}
                            titleClassName={styles.cardTitle}
                        >
                            {employee.aboutMe ? (
                                <p className={styles.bioText}>{employee.aboutMe}</p>
                            ) : (
                                <p className={styles.bioText} style={{ color: '#605e5c', fontStyle: 'italic' }}>No bio available</p>
                            )}
                        </ProfileSection>

                        <ProfileSection
                            icon="🎯"
                            title="Quick Summary"
                            cardClassName={styles.contentCard}
                            headerClassName={styles.cardHeader}
                            iconClassName={styles.cardIcon}
                            titleClassName={styles.cardTitle}
                        >
                            <div className={styles.contactGrid}>
                                <div className={styles.contactItem}>
                                    <div className={styles.contactLabel}>Current Role</div>
                                    <div className={styles.contactValue}>{employee.jobTitle || 'N/A'}</div>
                                </div>
                                <div className={styles.contactItem}>
                                    <div className={styles.contactLabel}>Department</div>
                                    <div className={styles.contactValue}>{employee.department}</div>
                                </div>
                            </div>
                        </ProfileSection>
                    </div>
                )}

                {activeTab === 'details' && (
                    <div className={`${styles.tabContent} ${styles.active}`}>
                        <div className={styles.detailsGrid} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '24px' }}>
                            <ProfileSection
                                icon="💡"
                                title="Core Skills"
                                cardClassName={styles.contentCard}
                                headerClassName={styles.cardHeader}
                                iconClassName={styles.cardIcon}
                                titleClassName={styles.cardTitle}
                            >
                                <div className={styles.skillsGrid} style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                    {employee.skills && employee.skills.length > 0 ? (
                                        employee.skills.map((skill, index) => (
                                            <div key={skill} className={styles.skillItem} style={{ padding: '4px 12px', background: '#f3f2f1', borderRadius: '16px', fontSize: '13px' }}>
                                                {skill}
                                            </div>
                                        ))
                                    ) : (
                                        <p style={{ color: '#605e5c', fontStyle: 'italic' }}>No skills listed</p>
                                    )}
                                </div>
                            </ProfileSection>

                            <ProfileSection
                                icon="🎯"
                                title="Interests"
                                cardClassName={styles.contentCard}
                                headerClassName={styles.cardHeader}
                                iconClassName={styles.cardIcon}
                                titleClassName={styles.cardTitle}
                            >
                                <div className={styles.skillsGrid} style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                    {employee.interests && employee.interests.length > 0 ? (
                                        employee.interests.map((interest, index) => (
                                            <div key={interest} className={styles.skillItem} style={{ padding: '4px 12px', background: '#fdf4f5', color: '#a4262c', borderRadius: '16px', fontSize: '13px' }}>
                                                {interest}
                                            </div>
                                        ))
                                    ) : (
                                        <p style={{ color: '#605e5c', fontStyle: 'italic' }}>No interests listed</p>
                                    )}
                                </div>
                            </ProfileSection>
                        </div>

                        <ProfileSection
                            icon="📁"
                            title="Recent Projects"
                            cardClassName={styles.contentCard}
                            headerClassName={styles.cardHeader}
                            iconClassName={styles.cardIcon}
                            titleClassName={styles.cardTitle}
                            style={{ marginTop: '24px' }}
                        >
                            <div className={styles.projectsList}>
                                {employee.pastProjects && employee.pastProjects.length > 0 ? (
                                    employee.pastProjects.map((project) => (
                                        <div key={project} className={styles.projectCard}>
                                            <div className={styles.projectIcon}>📱</div>
                                            <div className={styles.projectInfo}>
                                                <div className={styles.projectName}>{project}</div>
                                            </div>
                                        </div>
                                    ))
                                ) : (
                                    <p style={{ color: '#605e5c', fontSize: '14px' }}>No projects listed</p>
                                )}
                            </div>
                        </ProfileSection>
                    </div>
                )}

                {activeTab === 'organization' && (
                    <div className={`${styles.tabContent} ${styles.active}`}>
                        {/* People I Work With */}
                        <ProfileSection
                            icon="👥"
                            title="People I Work With"
                            cardClassName={styles.contentCard}
                            headerClassName={styles.cardHeader}
                            iconClassName={styles.cardIcon}
                            titleClassName={styles.cardTitle}
                        >
                            <div className={styles.colleaguesContainer}>
                                {employee.colleagues && employee.colleagues.length > 0 ? (
                                    employees
                                        .filter((emp: IEmployee) => employee.colleagues?.includes(emp.id))
                                        .map((colleague: IEmployee) => (
                                            <button
                                                key={colleague.id}
                                                className={styles.colleagueAvatar}
                                                title={colleague.displayName}
                                                onClick={() => {
                                                    if (onAuditLog) onAuditLog('Profile Interaction', colleague.displayName || colleague.mail || '', { source: 'ColleagueCircle', action: 'ViewProfile' });
                                                    if (onEmployeeSelect) onEmployeeSelect(colleague);
                                                }}
                                                aria-label={`View profile of ${colleague.displayName}`}
                                            >
                                                {colleague.initials}
                                            </button>
                                        ))
                                ) : (
                                    <p style={{ color: '#605e5c', fontSize: '14px' }}>No colleagues listed</p>
                                )}
                            </div>
                        </ProfileSection>

                        {/* Organization Structure */}
                        <ProfileOrgChart
                            employee={employee}
                            employees={employees}
                            onEmployeeSelect={onEmployeeSelect}
                            onAuditLog={onAuditLog}
                            orgChartLayout={orgChartLayout}
                            styles={styles}
                            cardClassName={styles.contentCard}
                        />
                    </div>
                )}
            </div>
        </div>
    );
};
