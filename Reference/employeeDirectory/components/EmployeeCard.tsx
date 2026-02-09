import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployee } from '../models/IEmployee';
import { GraphService } from '../services/GraphService';
import { Icon, PersonaPresence } from '@fluentui/react';

export interface IEmployeeCardProps {
    employee: IEmployee;
    onClick: (employee: IEmployee) => void;
    graphService: GraphService;
    presenceStatus?: string;
    animationDelay?: string;
    // layoutMode removed
    subHeadingFontSize?: number;
    contentFontSize?: number;
    kudosCounts: { [email: string]: number };
    onGiveKudos: (user: IEmployee) => void;
    onViewKudos: (user: IEmployee) => void;
    isKudosEnabled?: boolean;
    visualTheme?: 'modern-scrolling-premium' | 'modern-tabbed-premium';
    isListView?: boolean;
}

export const EmployeeCard: React.FunctionComponent<IEmployeeCardProps> = (props) => {
    const [photoUrl, setPhotoUrl] = React.useState<string>('');

    React.useEffect(() => {
        if (props.employee.id) {
            props.graphService.getUserPhoto(props.employee.id).then(url => {
                if (url) setPhotoUrl(url);
            });
        }
    }, [props.employee.id]);

    const _copyToClipboard = (text: string, e: React.MouseEvent | React.KeyboardEvent) => {
        e.stopPropagation();
        navigator?.clipboard?.writeText(text);
    };

    const _downloadVCard = (e: React.MouseEvent | React.KeyboardEvent) => {
        e.stopPropagation();
        const user = props.employee;
        const vCardData = `BEGIN:VCARD
VERSION:3.0
FN:${user.displayName}
N:;;;;
EMAIL:${user.mail || ''}
TEL;TYPE=CELL:${user.mobilePhone || ''}
TEL;TYPE=WORK:${user.businessPhones ? user.businessPhones[0] : ''}
TITLE:${user.jobTitle || ''}
ORG:${user.companyName || ''}
END:VCARD`;
        const blob = new Blob([vCardData], { type: 'text/vcard' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${user.displayName || 'contact'}.vcf`;
        document.body.appendChild(link);
        link.click();
        link.remove();
    };

    const _getPresence = (status: string): PersonaPresence => {
        switch (status) {
            case 'Available': return PersonaPresence.online;
            case 'Away': return PersonaPresence.away;
            case 'Busy': return PersonaPresence.busy;
            case 'DND': return PersonaPresence.dnd;
            case 'Offline': return PersonaPresence.offline;
            default: return PersonaPresence.none;
        }
    };

    const _renderPremiumCard = () => {
        const innerEmail = props.employee.mail || props.employee.userPrincipalName;
        const kudosCount = innerEmail ? (props.kudosCounts[innerEmail] || 0) : 0;
        const isListView = props.isListView;

        return (
            <button
                className={styles.employeeCard}
                onClick={() => props.onClick(props.employee)}
                style={{
                    border: 'none', padding: 0, font: 'inherit', cursor: 'pointer', textAlign: 'left',
                    animationDelay: props.animationDelay
                }}
            >
                {/* Left Column (3/4): Identity Info */}
                <div className={styles.cardContentLeft}>
                    <div className={styles.avatar}>
                        {photoUrl ? (
                            <img src={photoUrl} alt={props.employee.displayName} />
                        ) : (
                            <span>{props.employee.displayName?.charAt(0)}</span>
                        )}
                        <div className={styles.onlineStatus} style={{ background: props.presenceStatus === 'Available' ? 'var(--m-online-green)' : '#ccc' }}></div>
                    </div>

                    <div className={styles.employeeInfo}>
                        <div className={styles.employeeName}>{props.employee.displayName}</div>
                        <div className={styles.employeeTitle}>{props.employee.jobTitle}</div>

                        <div className={styles.employeeEmail}>
                            <Icon iconName="Mail" />
                            <span>{props.employee.mail}</span>
                        </div>

                        {props.employee.department && (
                            <div className={styles.employeeDept}>{props.employee.department}</div>
                        )}
                    </div>
                </div>

                {/* Right Column (1/4): Action List */}
                <div className={styles.cardContentRight} onClick={(e) => e.stopPropagation()}>
                    <div className={styles.verticalActionList}>
                        <div
                            className={styles.verticalActionItem}
                            onClick={() => globalThis.location.href = `mailto:${props.employee.mail}`}
                        >
                            <Icon iconName="Mail" />
                            <span>Send Email</span>
                        </div>
                        <div
                            className={styles.verticalActionItem}
                            onClick={() => globalThis.open(`https://teams.microsoft.com/l/chat/0/0?users=${props.employee.mail}`, '_blank')}
                        >
                            <Icon iconName="TeamsLogo" />
                            <span>Teams Chat</span>
                        </div>
                        {props.isKudosEnabled && (
                            <div
                                className={`${styles.verticalActionItem} ${styles.kudosAction}`}
                                onClick={() => props.onGiveKudos(props.employee)}
                            >
                                <Icon iconName="FavoriteStarFill" />
                                <span>Kudos {kudosCount > 0 ? `(${kudosCount})` : ''}</span>
                            </div>
                        )}
                    </div>
                </div>
            </button>
        );
    };

    return _renderPremiumCard();
};
