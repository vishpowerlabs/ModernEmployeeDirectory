import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';

export const SkeletonCard: React.FunctionComponent = () => {
    return (
        <div className={`${styles.employeeCard} ${styles.skeletonCard}`}>
            <div className={styles.cardHeader}>
                <div className={`${styles.skeletonPhoto} ${styles.shimmer}`}></div>
                <div className={styles.employeeInfo} style={{ width: '100%' }}>
                    <div className={`${styles.skeletonText} ${styles.shimmer}`} style={{ width: '60%', height: '16px', marginBottom: '8px' }}></div>
                    <div className={`${styles.skeletonText} ${styles.shimmer}`} style={{ width: '80%', height: '12px' }}></div>
                </div>
            </div>
            <div className={styles.cardDetails}>
                {[1, 2].map(i => (
                    <div className={styles.detailRow} key={i}>
                        <div className={`${styles.skeletonText} ${styles.shimmer}`} style={{ width: '20px', height: '20px', borderRadius: '50%' }}></div>
                        <div className={`${styles.skeletonText} ${styles.shimmer}`} style={{ flex: 1, height: '14px', marginLeft: '10px' }}></div>
                    </div>
                ))}
            </div>
            <div className={styles.cardActions}>
                <div className={`${styles.skeletonButton} ${styles.shimmer}`}></div>
                <div className={`${styles.skeletonButton} ${styles.shimmer}`}></div>
            </div>
        </div>
    );
};
