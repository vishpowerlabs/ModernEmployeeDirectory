import * as React from 'react';
import styles from './TopBar.module.scss';

export interface ITopBarProps {
    onOpenPanel: (mode: 'profile' | 'kudos' | 'updateProfile') => void;
    title?: string;
    subtitle?: string;
    currentUserInitials?: string;
    buildNumber?: string;
}

export const TopBar: React.FunctionComponent<ITopBarProps> = (props) => {
    const { onOpenPanel, title = 'Employee Directory', subtitle = 'Find and connect with colleagues', currentUserInitials = 'CU', buildNumber = '0.0.1' } = props;
    const [isMenuOpen, setIsMenuOpen] = React.useState(false);
    const menuRef = React.useRef<HTMLDivElement>(null);

    // Close menu when clicking outside
    React.useEffect(() => {
        const handleClickOutside = (event: MouseEvent): void => {
            if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
                setIsMenuOpen(false);
            }
        };

        if (isMenuOpen) {
            document.addEventListener('mousedown', handleClickOutside);
        }

        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };
    }, [isMenuOpen]);

    const handleMenuAction = (mode: 'profile' | 'kudos' | 'updateProfile'): void => {
        onOpenPanel(mode);
        setIsMenuOpen(false);
    };

    return (
        <div className={styles.header}>
            <div className="container">
                <div className={styles.headerInner}>
                    <div className={styles.headerContent}>
                        <h1>{title}</h1>
                        <p>{subtitle}</p>
                    </div>

                    <div className={styles.headerActions} ref={menuRef}>
                        <button
                            className={styles.hamburgerBtn}
                            onClick={() => setIsMenuOpen(!isMenuOpen)}
                            title="Menu"
                            type="button"
                            aria-expanded={isMenuOpen}
                        >
                            <span className={styles.hamburgerIcon}></span>
                        </button>

                        {isMenuOpen && (
                            <div className={styles.dropdownMenu}>
                                <div className={styles.menuHeader}>
                                    <div className={styles.userAvatar}>{currentUserInitials}</div>
                                    <div className={styles.userInfo}>
                                        <span className={styles.userName}>My Account</span>
                                    </div>
                                </div>
                                <div className={styles.menuDivider}></div>
                                <button className={styles.menuItem} onClick={() => handleMenuAction('profile')}>
                                    <span className={styles.menuIcon}>👤</span> My Profile
                                </button>
                                <button className={styles.menuItem} onClick={() => handleMenuAction('kudos')}>
                                    <span className={styles.menuIcon}>⭐</span> My Kudos
                                </button>
                                <button className={styles.menuItem} onClick={() => handleMenuAction('updateProfile')}>
                                    <span className={styles.menuIcon}>✏️</span> Update Profile
                                </button>
                                <div className={styles.menuDivider}></div>
                                <div className={styles.menuFooter}>
                                    <span className={styles.buildLabel}>Build v{buildNumber}</span>
                                </div>
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </div>
    );
};
