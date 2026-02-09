import * as React from 'react';
import styles from './ProfileTabbed.module.scss';
import { IEmployee } from '../Home/DirectoryHome';
import { OrgChart } from '../Shared/OrgChart';

export interface IProfileTabbedProps {
    employee: IEmployee;
    employees: IEmployee[];  // All employees for org chart
    onBack: () => void;
    onKudosClick: () => void;
    kudosCount?: number;
    onEmployeeSelect?: (employee: IEmployee) => void;  // Callback when clicking employee in org chart
    orgChartLayout?: 'vertical' | 'horizontal' | 'compact';
}

export const ProfileTabbed: React.FunctionComponent<IProfileTabbedProps> = (props) => {
    const { employee, employees, onBack, kudosCount = 0, onEmployeeSelect, orgChartLayout } = props;
    const [activeTab, setActiveTab] = React.useState('overview');

    return (
        <div className={styles.profileTabbed}>
            {/* Top Bar */}
            <div className={styles.topBar}>
                <div className={styles.topBarInner}>
                    <div className={styles.pageTitle}>
                        👥 Employee Directory
                    </div>
                    <button className={styles.backBtn} onClick={onBack}>
                        ← Back to Directory
                    </button>
                </div>
            </div>

            <div className={styles.container}>
                {/* Profile Header */}
                <div className={styles.profileHeader}>
                    <div className={styles.profileLeft}>
                        <div className={styles.avatarLarge}>
                            {employee.initials}
                            {employee.isOnline && <div className={styles.onlineStatus}></div>}
                        </div>
                        <div className={styles.profileInfo}>
                            <h1>{employee.displayName}</h1>
                            <p className={styles.profileTitle}>{employee.jobTitle || 'N/A'}</p>
                            <div className={styles.profileLocation}>
                                <span className={styles.headerContactItem}>✉️ {employee.mail || employee.userPrincipalName || 'N/A'}</span>
                                <span className={styles.headerContactItem}>📞 {employee.businessPhones?.[0] || employee.mobilePhone || 'N/A'}</span>
                                <span className={styles.headerContactItem}>📍 {employee.city && employee.state ? `${employee.city}, ${employee.state}` : employee.officeLocation || 'N/A'}</span>
                            </div>
                        </div>
                    </div>
                    <div className={styles.quickActions}>
                        <button className={`${styles.actionBtn} ${styles.primary}`}>✉️ Email</button>
                        <button className={styles.actionBtn}>💬 Teams</button>
                        <button className={styles.actionBtn}>📞 Call</button>
                        <button className={styles.btnKudos} onClick={props.onKudosClick}>
                            ⭐ Kudos {kudosCount > 0 && <span className={styles.kudosCount}>{kudosCount}</span>}
                        </button>
                    </div>
                </div>

                {/* Tab Navigation */}
                <div className={styles.tabNavigation}>
                    <button
                        className={`${styles.tabBtn} ${activeTab === 'overview' ? styles.active : ''}`}
                        onClick={() => setActiveTab('overview')}
                    >
                        📋 Overview
                    </button>

                    <button
                        className={`${styles.tabBtn} ${activeTab === 'details' ? styles.active : ''}`}
                        onClick={() => setActiveTab('details')}
                    >
                        📝 More Details
                    </button>
                    <button
                        className={`${styles.tabBtn} ${activeTab === 'organization' ? styles.active : ''}`}
                        onClick={() => setActiveTab('organization')}
                    >
                        🏢 Organization
                    </button>
                </div>

                {/* Tab Content */}
                {activeTab === 'overview' && (
                    <div className={`${styles.tabContent} ${styles.active}`}>
                        <div className={styles.contentCard}>
                            <div className={styles.cardHeader}>
                                <div className={styles.cardIcon}>👤</div>
                                <h2 className={styles.cardTitle}>Professional Bio</h2>
                            </div>
                            {employee.aboutMe ? (
                                <p className={styles.bioText}>{employee.aboutMe}</p>
                            ) : (
                                <p className={styles.bioText} style={{ color: '#605e5c', fontStyle: 'italic' }}>No bio available</p>
                            )}
                        </div>
                        <div className={styles.contentCard}>
                            <div className={styles.cardHeader}>
                                <div className={styles.cardIcon}>🎯</div>
                                <h2 className={styles.cardTitle}>Quick Summary</h2>
                            </div>
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
                        </div>
                    </div>
                )}



                {activeTab === 'details' && (
                    <div className={`${styles.tabContent} ${styles.active}`}>
                        <div className={styles.detailsGrid} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '24px' }}>
                            <div className={styles.contentCard}>
                                <div className={styles.cardHeader}>
                                    <div className={styles.cardIcon}>💡</div>
                                    <h2 className={styles.cardTitle}>Core Skills</h2>
                                </div>
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
                            </div>

                            <div className={styles.contentCard}>
                                <div className={styles.cardHeader}>
                                    <div className={styles.cardIcon}>🎯</div>
                                    <h2 className={styles.cardTitle}>Interests</h2>
                                </div>
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
                            </div>
                        </div>

                        <div className={styles.contentCard} style={{ marginTop: '24px' }}>
                            <div className={styles.cardHeader}>
                                <div className={styles.cardIcon}>📁</div>
                                <h2 className={styles.cardTitle}>Recent Projects</h2>
                            </div>
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
                        </div>
                    </div>
                )}

                {activeTab === 'organization' && (
                    <div className={`${styles.tabContent} ${styles.active}`}>
                        {/* People I Work With */}
                        <div className={styles.contentCard}>
                            <div className={styles.cardHeader}>
                                <div className={styles.cardIcon}>👥</div>
                                <h2 className={styles.cardTitle}>People I Work With</h2>
                            </div>
                            <div className={styles.colleaguesContainer}>
                                {employee.colleagues && employee.colleagues.length > 0 ? (
                                    employees
                                        .filter((emp: IEmployee) => employee.colleagues?.includes(emp.id))
                                        .map((colleague: IEmployee) => (
                                            <div
                                                key={colleague.id}
                                                className={styles.colleagueAvatar}
                                                title={colleague.displayName}
                                            >
                                                {colleague.initials}
                                            </div>
                                        ))
                                ) : (
                                    <p style={{ color: '#605e5c', fontSize: '14px' }}>No colleagues listed</p>
                                )}
                            </div>
                        </div>

                        {/* Organization Structure */}
                        <div className={styles.contentCard}>
                            <div className={styles.cardHeader}>
                                <div className={styles.cardIcon}>🏢</div>
                                <h2 className={styles.cardTitle}>Organization Structure</h2>
                            </div>
                            <OrgChart
                                employees={employees}
                                currentEmployeeId={employee.id}
                                currentEmployee={employee}
                                onEmployeeClick={onEmployeeSelect}
                                layout={orgChartLayout}
                            />
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};
