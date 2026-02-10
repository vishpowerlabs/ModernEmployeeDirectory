import * as React from 'react';
import { IEmployee } from '../Home/DirectoryHome';
import { OrgChart } from '../Shared/OrgChart';

export interface IProfileCommonProps {
    employee: IEmployee;
    onContactClick: (type: 'Email' | 'Teams' | 'Call') => void;
    onKudosClick?: () => void;
    kudosCount?: number;
    styles: any;
}

export const ContactItem: React.FC<{ icon: string; text: string; className: string }> = ({ icon, text, className }) => (
    <span className={className}>{icon} {text}</span>
);

export const ActionButton: React.FC<{
    icon: string;
    label: string;
    onClick: () => void;
    className: string;
    variant?: 'primary' | 'secondary' | 'kudos';
    kudosCount?: number;
    kudosCountClassName?: string;
}> = ({ icon, label, onClick, className, variant, kudosCount, kudosCountClassName }) => (
    <button className={className} onClick={onClick}>
        {icon} {label}
        {variant === 'kudos' && kudosCount !== undefined && kudosCount > 0 && (
            <span className={kudosCountClassName}>{kudosCount}</span>
        )}
    </button>
);

export const ProfileSection: React.FC<{
    icon: string;
    title: string;
    children: React.ReactNode;
    headerClassName: string;
    iconClassName: string;
    titleClassName: string;
    cardClassName: string;
    style?: React.CSSProperties;
}> = ({ icon, title, children, headerClassName, iconClassName, titleClassName, cardClassName, style }) => (
    <div className={cardClassName} style={style}>
        <div className={headerClassName}>
            <div className={iconClassName}>{icon}</div>
            <h3 className={titleClassName}>{title}</h3>
        </div>
        {children}
    </div>
);

/**
 * Shared Top Bar for Profile views
 */
export const ProfileTopBar: React.FC<{
    onBack: () => void;
    styles: any;
}> = ({ onBack, styles }) => (
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
);

/**
 * Shared Org Chart section for Profile views with consistent auditing
 */
export const ProfileOrgChart: React.FC<{
    employee: IEmployee;
    employees: IEmployee[];
    onEmployeeSelect?: (employee: IEmployee) => void;
    onAuditLog?: (activity: string, target: string, details: any) => void;
    orgChartLayout?: 'vertical' | 'horizontal' | 'compact';
    styles: any;
    cardClassName?: string;
    headerClassName?: string;
    iconClassName?: string;
    titleClassName?: string;
}> = ({ employee, employees, onEmployeeSelect, onAuditLog, orgChartLayout, styles, cardClassName, headerClassName, iconClassName, titleClassName }) => (
    <ProfileSection
        icon="🏢"
        title="Organization Structure"
        cardClassName={cardClassName || `${styles.sectionCard} ${styles.fullWidth}`}
        headerClassName={headerClassName || styles.cardHeader}
        iconClassName={iconClassName || styles.cardIcon}
        titleClassName={titleClassName || styles.cardTitle}
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
);

/**
 * Shared helper for contact click actions (Email, Teams, Call) with auditing
 */
export const handleProfileContactClick = (
    type: 'Email' | 'Teams' | 'Call',
    employee: IEmployee,
    onAuditLog?: (activity: string, target: string, details: any) => void
): void => {
    if (onAuditLog) {
        onAuditLog(
            `Profile Contact: ${type}`,
            employee.displayName || employee.mail || employee.id,
            {
                contactType: type,
                targetMail: employee.mail,
                functionName: 'handleProfileContactClick'
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
