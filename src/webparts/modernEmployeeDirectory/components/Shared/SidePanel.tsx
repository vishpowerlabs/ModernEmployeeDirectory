import * as React from 'react';
import styles from './SidePanel.module.scss';
import GlobalStyles from '../ModernEmployeeDirectory.module.scss';
import { IEmployee } from '../Home/DirectoryHome';
import { KudosFeed } from '../Kudos/KudosFeed';
import { GiveKudosForm } from '../Kudos/GiveKudosForm';
import { IKudos } from '../../models/IKudos';
import { UpdateProfileForm } from '../Profile/UpdateProfileForm';

export interface ISidePanelProps {
    isOpen: boolean;
    onClose: () => void;
    panelMode: 'profile' | 'kudos' | 'updateProfile';
    onSwitchToKudos: () => void;
    onSwitchToUpdateProfile: () => void;
    user: {
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
        pastProjects?: string[];
    };
    // Update Profile props
    updatableFields: string[];
    onUpdateProfile: (updatedFields: any) => Promise<boolean>;
    // Kudos props
    enableKudos: boolean;
    selectedEmployee?: IEmployee | null;
    kudos: IKudos[];
    kudosLoading: boolean;
    badgeChoices: string[];
    onGiveKudos: (message: string, badgeType: string) => Promise<void>;
    loading?: boolean;
}

export const SidePanel: React.FunctionComponent<ISidePanelProps> = (props) => {
    const {
        isOpen, onClose, panelMode, onSwitchToKudos, onSwitchToUpdateProfile,
        user, enableKudos, selectedEmployee, kudos, kudosLoading,
        onGiveKudos, updatableFields, onUpdateProfile, loading, badgeChoices
    } = props;
    const [showGiveKudosForm, setShowGiveKudosForm] = React.useState(false);

    // Add ESC key listener when panel is open
    React.useEffect(() => {
        const handleEscKey = (event: KeyboardEvent): void => {
            if ((event.key === 'Escape' || event.key === 'Esc') && isOpen) {
                event.preventDefault();
                event.stopPropagation();
                onClose();
            }
        };

        if (isOpen) {
            globalThis.addEventListener('keydown', handleEscKey, { capture: true });
        }

        return () => {
            globalThis.removeEventListener('keydown', handleEscKey, { capture: true });
        };
    }, [isOpen, onClose]);

    // Reset form state when panel mode changes or closes
    React.useEffect(() => {
        setShowGiveKudosForm(false);
    }, [panelMode, isOpen]);

    const handleGiveKudosSubmit = async (message: string, badgeType: string): Promise<void> => {
        await onGiveKudos(message, badgeType);
        setShowGiveKudosForm(false);
    };

    const renderHeaderTitle = (): string => {
        if (panelMode === 'kudos') return 'Kudos';
        if (panelMode === 'updateProfile') return 'Update Profile';
        return 'My Profile';
    };

    const renderProfile = (): JSX.Element => (
        <>
            <div className={styles.userProfileCard}>
                <div className={styles.userProfileAvatar}>
                    {user.photoUrl ? (
                        <img src={user.photoUrl} alt={user.name} className={styles.avatarImage} />
                    ) : (
                        user.initials
                    )}
                    <div className={styles.userOnlineStatus}></div>
                </div>
                <h3 className={styles.userProfileName}>{user.name}</h3>
            </div>

            <div className={styles.userPanelActions}>
                {enableKudos && (
                    <button className={`${styles.userActionBtn} ${styles.primary}`} onClick={onSwitchToKudos}>
                        <span>⭐</span> View My Kudos
                    </button>
                )}
                <button className={styles.userActionBtn} onClick={onSwitchToUpdateProfile}>
                    <span>✏️</span> Update Profile
                </button>
            </div>
        </>
    );

    const renderKudos = (): JSX.Element => {
        if (showGiveKudosForm && selectedEmployee) {
            return (
                <GiveKudosForm
                    recipientName={selectedEmployee.displayName}
                    badgeChoices={badgeChoices}
                    onSubmit={handleGiveKudosSubmit}
                    onCancel={() => setShowGiveKudosForm(false)}
                />
            );
        }

        if (selectedEmployee?.mail) {
            return (
                <KudosFeed
                    kudos={kudos}
                    loading={kudosLoading}
                    emptyMessage={selectedEmployee.mail === user.email ? "You haven't received any kudos yet." : "No kudos found for this employee."}
                    onGiveKudos={() => setShowGiveKudosForm(true)}
                />
            );
        }

        if (selectedEmployee) {
            return (
                <div className={styles.emptyKudos}>
                    <p>Kudos system is restricted to employees with email addresses.</p>
                </div>
            );
        }

        return <></>;
    };

    const renderContent = (): JSX.Element => {
        switch (panelMode) {
            case 'updateProfile':
                return (
                    <UpdateProfileForm
                        user={user}
                        updatableFields={updatableFields}
                        onSubmit={onUpdateProfile}
                        onCancel={onClose}
                        loading={loading}
                    />
                );
            case 'kudos':
                return renderKudos();
            default:
                return renderProfile();
        }
    };

    return (
        <>
            {isOpen && (
                <button
                    className={`${GlobalStyles.overlay} ${isOpen ? 'open' : ''}`}
                    onClick={onClose}
                    onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}
                    aria-label="Close panel"
                    type="button"
                />
            )}
            <div className={`${styles.userPanel} ${isOpen ? styles.open : ''}`}>
                <div className={styles.userPanelHeader}>
                    <h2>{renderHeaderTitle()}</h2>
                    <button className={styles.closePanel} onClick={onClose}>✕</button>
                </div>
                <div className={styles.userPanelContent}>
                    {renderContent()}
                </div>
            </div>
        </>
    );
};
