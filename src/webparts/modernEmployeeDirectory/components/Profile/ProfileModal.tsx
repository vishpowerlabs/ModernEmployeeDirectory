import * as React from 'react';
import { useEffect, useRef } from 'react';
import styles from './ProfileModal.module.scss';
import { ActionButton, ProfileOrgChart, handleProfileContactClick } from './ProfileCommon';
import { IEmployee } from '../Home/DirectoryHome';

export interface IProfileModalProps {
    employee: IEmployee;
    employees: IEmployee[];
    kudosCount?: number;
    onClose: () => void;
    onKudosClick: () => void;
    onEmployeeSelect: (employee: IEmployee) => void;
    orgChartLayout?: 'vertical' | 'horizontal' | 'compact';
    onAuditLog?: (activity: string, target: string, details: any) => void;
}

export const ProfileModal: React.FC<IProfileModalProps> = ({
    employee,
    employees,
    kudosCount = 0,
    onClose,
    onKudosClick,
    onEmployeeSelect,
    orgChartLayout = 'vertical',
    onAuditLog
}) => {
    const modalRef = useRef<HTMLDivElement>(null);

    // Handle ESC key to close
    useEffect(() => {
        const handleEsc = (event: KeyboardEvent) => {
            if (event.key === 'Escape') {
                onClose();
            }
        };
        window.addEventListener('keydown', handleEsc);
        return () => window.removeEventListener('keydown', handleEsc);
    }, [onClose]);

    // Handle backdrop click
    const handleBackdropClick = (e: React.MouseEvent) => {
        if (e.target === e.currentTarget) {
            onClose();
        }
    };

    const handleContact = (type: 'Email' | 'Teams' | 'Call') => {
        handleProfileContactClick(type, employee, onAuditLog);
    };

    return (
        <div className={styles.profileModal} onClick={handleBackdropClick} role="dialog" aria-modal="true">
            <div className={styles.modalCard} ref={modalRef}>
                <button className={styles.closeButton} onClick={onClose} aria-label="Close Profile">
                    ✕
                </button>

                {/* Left Sidebar */}
                <div className={styles.sidebar}>
                    <div className={styles.photoContainer}>
                        {employee.photoUrl ? (
                            <img src={employee.photoUrl} alt={employee.displayName} className={styles.profileImage} />
                        ) : (
                            <div className={styles.initials}>{employee.initials}</div>
                        )}
                    </div>

                    <h2 className={styles.name}>{employee.displayName}</h2>
                    <div className={styles.title}>{employee.jobTitle}</div>

                    <div className={styles.contactGrid}>
                        <button className={styles.contactBtn} onClick={() => handleContact('Email')} title={employee.mail}>
                            📧 Email
                        </button>
                        <button className={styles.contactBtn} onClick={() => handleContact('Teams')} title="Chat on Teams">
                            💬 Teams
                        </button>
                        <button className={styles.contactBtn} onClick={() => handleContact('Call')} title={employee.mobilePhone}>
                            📞 Call
                        </button>
                    </div>

                    <div className={styles.kudosButton}>
                        <ActionButton
                            icon="⭐"
                            label={`Give Kudos (${kudosCount})`}
                            onClick={onKudosClick}
                            className=""
                        />
                    </div>
                </div>

                {/* Main Content Area */}
                <div className={styles.content}>
                    <div className={styles.section}>
                        <h3>👤 About Me</h3>
                        <p>{employee.aboutMe || "No bio available."}</p>
                    </div>

                    {employee.skills && employee.skills.length > 0 && (
                        <div className={styles.section}>
                            <h3>🛠 Expertise</h3>
                            <div className={styles.tagContainer}>
                                {employee.skills.map((skill: string, i: number) => (
                                    <span key={i} className={styles.tag}>{skill}</span>
                                ))}
                            </div>
                        </div>
                    )}

                    {employee.interests && employee.interests.length > 0 && (
                        <div className={styles.section}>
                            <h3>💡 Interests</h3>
                            <div className={styles.tagContainer}>
                                {employee.interests.map((interest: string, i: number) => (
                                    <span key={i} className={styles.tag} style={{ background: '#107c10' }}>{interest}</span> // Green for interests
                                ))}
                            </div>
                        </div>
                    )}

                    {employee.projects && employee.projects.length > 0 && (
                        <div className={styles.section}>
                            <h3>🚀 Projects</h3>
                            <div className={styles.projectList}>
                                {employee.projects.map((project: string, i: number) => (
                                    <div key={i} className={styles.projectCard}>
                                        <strong>{project}</strong>
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}

                    <div className={styles.section}>
                        <ProfileOrgChart
                            employee={employee}
                            employees={employees}
                            onEmployeeSelect={onEmployeeSelect}
                            onAuditLog={onAuditLog}
                            orgChartLayout={orgChartLayout}
                            styles={styles}
                            cardClassName={styles.orgChartContainer} // Allow override if needed
                        />
                    </div>
                </div>
            </div>
        </div>
    );
};
