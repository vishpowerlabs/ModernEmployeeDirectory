import * as React from 'react';
import { IEmployee } from '../Home/DirectoryHome';

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
