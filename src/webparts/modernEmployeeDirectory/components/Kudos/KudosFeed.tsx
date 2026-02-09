import * as React from 'react';
import styles from './KudosFeed.module.scss';
import { IKudos } from '../../models/IKudos';

export interface IKudosFeedProps {
    kudos: IKudos[];
    loading: boolean;
    emptyMessage?: string;
    onGiveKudos: () => void;
}

const BADGE_ICONS: { [key: string]: string } = {
    star: '⭐',
    trophy: '🏆',
    heart: '❤️',
    lightbulb: '💡',
    rocket: '🚀',
    clap: '👏'
};

export const KudosFeed: React.FC<IKudosFeedProps> = ({ kudos, loading, emptyMessage, onGiveKudos }) => {
    const formatTimestamp = (date: Date): string => {
        const now = new Date();
        const diff = now.getTime() - date.getTime();
        const minutes = Math.floor(diff / 60000);
        const hours = Math.floor(diff / 3600000);
        const days = Math.floor(diff / 86400000);

        if (minutes < 60) return `${minutes} min ago`;
        if (hours < 24) return `${hours} hr ago`;
        if (days < 7) return `${days} day${days > 1 ? 's' : ''} ago`;
        return date.toLocaleDateString();
    };

    if (loading) {
        return (
            <div className={styles.kudosFeed}>
                <div className={styles.emptyState}>
                    <div className={styles.emptyText}>Loading kudos...</div>
                </div>
            </div>
        );
    }

    return (
        <div className={styles.kudosFeed}>
            <div className={styles.header}>
                <h2 className={styles.title}>Received Kudos</h2>
                <button className={styles.giveKudosButton} onClick={onGiveKudos}>
                    ➕ Give Kudos
                </button>
            </div>

            {kudos.length === 0 ? (
                <div className={styles.emptyState}>
                    <div className={styles.emptyIcon}>🌟</div>
                    <div className={styles.emptyText}>{emptyMessage || 'No kudos received yet'}</div>
                </div>
            ) : (
                kudos.map((kudo) => (
                    <div key={kudo.id} className={styles.kudosCard}>
                        <div className={styles.cardHeader}>
                            <div className={styles.badgeIcon}>
                                {BADGE_ICONS[kudo.badgeType.toLowerCase()] || '⭐'}
                            </div>
                            <div className={styles.authorInfo}>
                                <div className={styles.authorName}>{kudo.authorName}</div>
                                <div className={styles.timestamp}>{formatTimestamp(kudo.createdDate)}</div>
                            </div>
                        </div>
                        <div className={styles.message}>{kudo.message}</div>
                    </div>
                ))
            )}
        </div>
    );
};
