import * as React from 'react';
import styles from './ProfileScrolling.module.scss';
import { IEmployee } from '../Home/DirectoryHome';
import { OrgChart } from '../Shared/OrgChart';
import { ContactItem, ActionButton, ProfileSection } from './ProfileCommon';

export interface IProfileScrollingProps {
    employee: IEmployee;
    employees: IEmployee[];  // All employees for org chart
    onBack: () => void;
    onKudosClick: () => void;
    kudosCount?: number;
    onEmployeeSelect?: (employee: IEmployee) => void;  // Callback when clicking employee in org chart
    orgChartLayout?: 'vertical' | 'horizontal' | 'compact';
    onAuditLog?: (activity: string, target: string, details: any) => void;
}

export const ProfileScrolling: React.FunctionComponent<IProfileScrollingProps> = (props) => {
    const { employee, employees, onBack, onKudosClick, kudosCount = 0, onEmployeeSelect, orgChartLayout, onAuditLog } = props;

    const _handleContactClick = (type: 'Email' | 'Teams' | 'Call'): void => {
        if (onAuditLog) {
            onAuditLog(
                `Profile Contact: ${type}`,
                employee.displayName || employee.mail || employee.id,
                {
                    contactType: type,
                    targetMail: employee.mail,
                    functionName: '_handleContactClick'
                }
            );
        }

        switch (type) {
            case 'Email':
                if (employee.mail) globalThis.location.href = `mailto:${employee.mail}`;
                break;
            case 'Teams':
                if (employee.mail) globalThis.open(`https://teams.microsoft.com/l/chat/0/0?users=${employee.mail}`, '_blank');
                break;
            case 'Call': {
                const phone = employee.businessPhones?.[0] || employee.mobilePhone;
                if (phone) globalThis.location.href = `tel:${phone}`;
                break;
            }
        }
    };

    return (
        <div className={styles.profileScrolling}>
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
                    <div className={styles.profilePhoto}>
                        <div className={styles.avatarLarge}>
                            {employee.photoUrl ? (
                                <img src={employee.photoUrl} alt={employee.displayName} style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover' }} />
                            ) : (
                                employee.initials
                            )}
                            {employee.isOnline && <div className={styles.onlineStatus}></div>}
                        </div>
                    </div>
                    <div className={styles.profileInfo}>
                        <h1 className={styles.profileName}>{employee.displayName}</h1>
                        <div className={styles.profileTitle}>{employee.jobTitle || 'N/A'}</div>
                        <div className={styles.profileLocation}>
                            <ContactItem icon="✉️" text={employee.mail || employee.userPrincipalName || 'N/A'} className={styles.headerContactItem} />
                            <ContactItem icon="📞" text={employee.businessPhones?.[0] || employee.mobilePhone || 'N/A'} className={styles.headerContactItem} />
                            <ContactItem icon="📍" text={employee.city && employee.state ? `${employee.city}, ${employee.state}` : employee.officeLocation || 'N/A'} className={styles.headerContactItem} />
                        </div>

                        <div className={styles.actionButtons}>
                            <ActionButton icon="✉️" label="Email" onClick={() => _handleContactClick('Email')} className={styles.btnPrimary} />
                            <ActionButton icon="💬" label="Teams" onClick={() => _handleContactClick('Teams')} className={styles.btnSecondary} />
                            <ActionButton icon="📞" label="Call" onClick={() => _handleContactClick('Call')} className={styles.btnSecondary} />
                            <ActionButton
                                icon="⭐"
                                label="Give Kudos"
                                onClick={() => {
                                    if (onAuditLog) onAuditLog('Kudos Interaction', employee.displayName || employee.mail || '', { source: 'ProfileButton', action: 'OpenKudosPanel' });
                                    onKudosClick();
                                }}
                                className={styles.btnKudos}
                                variant="kudos"
                                kudosCount={kudosCount}
                                kudosCountClassName={styles.kudosCount}
                            />
                        </div>
                    </div>
                </div>

                {/* Content Grid */}
                <div className={styles.contentGrid}>
                    <ProfileSection
                        icon="👤"
                        title="Professional Bio"
                        cardClassName={`${styles.sectionCard} ${styles.fullWidth}`}
                        headerClassName={styles.cardHeader}
                        iconClassName={styles.cardIcon}
                        titleClassName={styles.cardTitle}
                    >
                        <div className={styles.cardContent}>
                            {employee.aboutMe ? (
                                <p>{employee.aboutMe}</p>
                            ) : (
                                <p style={{ color: '#605e5c', fontSize: '14px' }}>No bio available</p>
                            )}
                        </div>
                    </ProfileSection>

                    {/* Skills */}
                    <ProfileSection
                        icon="💡"
                        title="Skills & Expertise"
                        cardClassName={styles.sectionCard}
                        headerClassName={styles.cardHeader}
                        iconClassName={styles.cardIcon}
                        titleClassName={styles.cardTitle}
                    >
                        <div className={styles.skillsContainer}>
                            {employee.skills && employee.skills.length > 0 ? (
                                employee.skills.map((skill) => (
                                    <span key={skill} className={styles.skillTag}>{skill}</span>
                                ))
                            ) : (
                                <p style={{ color: '#605e5c', fontSize: '14px' }}>No skills listed</p>
                            )}
                        </div>
                    </ProfileSection>

                    {/* Interests */}
                    <ProfileSection
                        icon="🎯"
                        title="Interests"
                        cardClassName={styles.sectionCard}
                        headerClassName={styles.cardHeader}
                        iconClassName={styles.cardIcon}
                        titleClassName={styles.cardTitle}
                    >
                        <div className={styles.skillsContainer}>
                            {employee.interests && employee.interests.length > 0 ? (
                                employee.interests.map((interest) => (
                                    <span key={interest} className={styles.skillTag} style={{ background: '#fdf4f5', color: '#a4262c' }}>{interest}</span>
                                ))
                            ) : (
                                <p style={{ color: '#605e5c', fontSize: '14px' }}>No interests listed</p>
                            )}
                        </div>
                    </ProfileSection>

                    {/* Projects */}
                    <ProfileSection
                        icon="📁"
                        title="Recent Projects"
                        cardClassName={styles.sectionCard}
                        headerClassName={styles.cardHeader}
                        iconClassName={styles.cardIcon}
                        titleClassName={styles.cardTitle}
                    >
                        <div className={styles.projectsList}>
                            {employee.pastProjects && employee.pastProjects.length > 0 ? (
                                employee.pastProjects.map((project) => (
                                    <div key={project} className={styles.projectItem}>
                                        <div className={styles.projectIcon}></div>
                                        <div>
                                            <div style={{ fontSize: '14px', fontWeight: 600 }}>{project}</div>
                                        </div>
                                    </div>
                                ))
                            ) : (
                                <p style={{ color: '#605e5c', fontSize: '14px' }}>No projects listed</p>
                            )}
                        </div>
                    </ProfileSection>

                    {/* Organization Chart */}
                    <ProfileSection
                        icon="🏢"
                        title="Organization Structure"
                        cardClassName={`${styles.sectionCard} ${styles.fullWidth}`}
                        headerClassName={styles.cardHeader}
                        iconClassName={styles.cardIcon}
                        titleClassName={styles.cardTitle}
                    >
                        <OrgChart
                            employees={employees}
                            currentEmployeeId={employee.id}
                            currentEmployee={employee}
                            onEmployeeClick={(emp) => {
                                if (onAuditLog) onAuditLog('Org Chart Interaction', emp.displayName || emp.mail || '', { source: 'OrgChartNode', action: 'ViewProfile' });
                                if (onEmployeeSelect) onEmployeeSelect(emp);
                            }}
                            layout={orgChartLayout}
                        />
                    </ProfileSection>
                </div>
            </div>
        </div>
    );
};
