import * as React from 'react';
import styles from './OrgChart.module.scss';
import { IEmployee } from '../models/IEmployee';
import { GraphService } from '../services/GraphService';

export interface IOrgChartProps {
    currentUser: IEmployee;
    graphService: GraphService;
    onUserSelect: (employee: IEmployee) => void;
}

export const OrgChart: React.FunctionComponent<IOrgChartProps> = (props) => {
    const [manager, setManager] = React.useState<IEmployee | null>(null);
    const [directReports, setDirectReports] = React.useState<IEmployee[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);

    React.useEffect(() => {
        let isMounted = true;
        const fetchOrgData = async () => {
            setLoading(true);
            setManager(null);
            setDirectReports([]);

            try {
                // Fetch Manager
                const fetchedManager = await props.graphService.getManager(props.currentUser.id);
                if (isMounted && fetchedManager) setManager(fetchedManager);

                // Fetch Direct Reports
                const fetchedReports = await props.graphService.getDirectReports(props.currentUser.id);
                if (isMounted) setDirectReports(fetchedReports || []);

            } catch (error) {
                console.error("Error loading org chart", error);
            } finally {
                if (isMounted) setLoading(false);
            }
        };

        if (props.currentUser?.id) {
            fetchOrgData();
        }

        return () => { isMounted = false; };
    }, [props.currentUser?.id]);

    const renderNode = (employee: IEmployee, type: 'manager' | 'current' | 'report') => {
        const initials = employee.displayName
            ? employee.displayName.split(' ').map((n) => n[0]).join('').substring(0, 2).toUpperCase()
            : '??';

        return (
            <button
                type="button"
                className={`${styles.treeNode} ${type === 'current' ? styles.currentUser : ''}`}
                onClick={() => type !== 'current' && props.onUserSelect(employee)}
                title={type === 'current' ? 'Current Selection' : 'Click to view profile'}
                disabled={type === 'current'}
            >
                <div className={styles.nodeContent}>
                    <div className={styles.nodeAvatar}>
                        {/* Note: We could load photos here too, but staying simple for now */}
                        {initials}
                    </div>
                    <div className={styles.nodeName}>{employee.displayName}</div>
                    <div className={styles.nodeTitle}>{employee.jobTitle}</div>
                </div>
            </button>
        );
    };

    if (loading) {
        return <div className={styles.orgChartContainer}>Loading Organization Chart...</div>;
    }

    if (!manager && directReports.length === 0) {
        return null; // Don't show empty org chart section
    }

    return (
        <div className={styles.orgChartContainer}>
            <div className={styles.orgChartTitle}>Organization Chart</div>

            {/* Manager Level */}
            {manager && (
                <div className={styles.managerSection}>
                    {renderNode(manager, 'manager')}
                </div>
            )}

            {/* Current User Level */}
            <div className={`${styles.currentSection} ${(directReports.length > 0) ? styles.hasReports : ''}`}>
                {renderNode(props.currentUser, 'current')}
            </div>

            {/* Direct Reports Level */}
            {directReports.length > 0 && (
                <div className={styles.reportsSection}>
                    {directReports.map(report => (
                        <div key={report.id} className={styles.reportNodeWrapper}>
                            {renderNode(report, 'report')}
                        </div>
                    ))}
                </div>
            )}
        </div>
    );
};
