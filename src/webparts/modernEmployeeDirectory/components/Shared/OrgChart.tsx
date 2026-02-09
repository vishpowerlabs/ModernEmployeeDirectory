import * as React from 'react';
import { Tree, TreeNode } from 'react-organizational-chart';
import { IEmployee } from '../Home/DirectoryHome';
import styles from './OrgChart.module.scss';

export interface IOrgChartProps {
    employees: IEmployee[];
    currentEmployeeId?: string;
    currentEmployee?: IEmployee; // Fallback for when the employee is not in the employees list
    onEmployeeClick?: (employee: IEmployee) => void;
    layout?: 'vertical' | 'horizontal' | 'compact';
}

export const OrgChart: React.FunctionComponent<IOrgChartProps> = (props) => {
    const { employees, currentEmployeeId, onEmployeeClick, layout = 'vertical' } = props;

    // Build employee map for quick lookup
    const employeeMap = new Map<string, IEmployee>();
    employees.forEach(emp => employeeMap.set(emp.id, emp));

    // Find current employee - using fallback if provided
    let currentEmployee = currentEmployeeId ? employeeMap.get(currentEmployeeId) : undefined;
    if (!currentEmployee && props.currentEmployee) {
        currentEmployee = props.currentEmployee;
        employeeMap.set(currentEmployee.id, currentEmployee);
    }

    if (!currentEmployee) {
        return <div className={styles.noData}>No employee selected</div>;
    }

    // Get manager (who they report to)
    const manager = currentEmployee.managerId ? employeeMap.get(currentEmployee.managerId) : undefined;

    // Get direct reports (who reports to them)
    const directReports = employees.filter(emp => emp.managerId === currentEmployee.id);

    // Get peers (who share the same manager)
    const peers = manager
        ? employees.filter(emp => emp.managerId === manager.id && emp.id !== currentEmployee.id)
        : [];

    // Helper to render a node (reused across tree and horizontal)
    const renderNode = (employee: IEmployee, isCurrent: boolean, relationLabel?: string): React.ReactElement => (
        <button
            className={`${styles.orgNode} ${isCurrent ? styles.current : ''}`}
            onClick={() => onEmployeeClick?.(employee)}
            style={{ cursor: onEmployeeClick ? 'pointer' : 'default' }}
            tabIndex={0}
        >
            <div
                className={styles.orgAvatar}
                style={{
                    background: isCurrent
                        ? 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'
                        : `linear-gradient(135deg, hsl(${((employee.id.codePointAt(0) || 0) * 137) % 360}, 70%, 60%) 0%, hsl(${((employee.id.codePointAt(0) || 0) * 137 + 60) % 360}, 70%, 50%) 100%)`
                }}
            >
                {employee.initials}
            </div>
            <div className={styles.orgInfo}>
                <div className={styles.orgName}>{employee.displayName}</div>
                <div className={styles.orgTitle}>{employee.jobTitle || 'N/A'}</div>
                {relationLabel && <div className={styles.orgRelation}>{relationLabel}</div>}
            </div>
        </button>
    );

    // Recursive function to build org chart tree for direct reports
    const buildOrgTree = (employee: IEmployee): React.ReactElement => {
        const reports = employees.filter(emp => emp.managerId === employee.id);
        const isCurrent = employee.id === currentEmployeeId;

        return (
            <TreeNode
                key={employee.id}
                label={renderNode(employee, isCurrent)}
            >
                {reports.map(report => buildOrgTree(report))}
            </TreeNode>
        );
    };

    // Compact List Layout
    if (layout === 'compact') {
        return (
            <div className={styles.compactContainer}>
                {manager && (
                    <button className={styles.compactRow} onClick={() => onEmployeeClick?.(manager)}>
                        <div className={styles.compactAvatar}>{manager.initials}</div>
                        <div className={styles.compactInfo}>
                            <div className={styles.compactName}>{manager.displayName}</div>
                            <div className={styles.compactTitle}>{manager.jobTitle || 'N/A'}</div>
                            <div className={styles.compactBadge}>Manager</div>
                        </div>
                    </button>
                )}
                <button className={`${styles.compactRow} ${styles.current}`} onClick={() => onEmployeeClick?.(currentEmployee)}>
                    <div className={styles.compactAvatar}>{currentEmployee.initials}</div>
                    <div className={styles.compactInfo}>
                        <div className={styles.compactName}>{currentEmployee.displayName}</div>
                        <div className={styles.compactTitle}>{currentEmployee.jobTitle || 'N/A'}</div>
                        <div className={styles.compactBadge}>Current</div>
                    </div>
                </button>
                {directReports.length > 0 && (
                    <div className={styles.compactReports}>
                        <div className={styles.compactLabel}>Direct Reports</div>
                        {directReports.map(report => (
                            <button key={report.id} className={styles.compactRow} onClick={() => onEmployeeClick?.(report)}>
                                <div className={styles.compactAvatar}>{report.initials}</div>
                                <div className={styles.compactInfo}>
                                    <div className={styles.compactName}>{report.displayName}</div>
                                    <div className={styles.compactTitle}>{report.jobTitle || 'N/A'}</div>
                                </div>
                            </button>
                        ))}
                    </div>
                )}
                {peers.length > 0 && (
                    <div className={styles.compactReports} style={{ borderLeftColor: 'rgba(0,0,0,0.05)' }}>
                        <div className={styles.compactLabel}>Works With (Peers)</div>
                        {peers.map(peer => (
                            <button key={peer.id} className={styles.compactRow} onClick={() => onEmployeeClick?.(peer)}>
                                <div className={styles.compactAvatar}>{peer.initials}</div>
                                <div className={styles.compactInfo}>
                                    <div className={styles.compactName}>{peer.displayName}</div>
                                    <div className={styles.compactTitle}>{peer.jobTitle || 'N/A'}</div>
                                </div>
                            </button>
                        ))}
                    </div>
                )}
            </div>
        );
    }

    // Custom Horizontal Flow Layout
    if (layout === 'horizontal') {
        return (
            <div className={styles.orgChartContainer}>
                <div className={styles.horizontalGrid}>
                    {/* Header Row for Labels */}
                    <div className={styles.gridHeader}>{manager && <div className={styles.compactLabel}>Manager</div>}</div>
                    <div className={styles.gridEmpty}></div>
                    <div className={styles.gridHeader}><div className={styles.compactLabel}>Current</div></div>
                    <div className={styles.gridEmpty}></div>
                    <div className={styles.gridHeader}>{directReports.length > 0 && <div className={styles.compactLabel}>Direct Reports</div>}</div>

                    {/* Content Row for Cards and Connectors */}
                    <div className={styles.gridCell}>
                        {manager && renderNode(manager, false)}
                    </div>

                    <div className={styles.gridCell}>
                        {manager && <div className={styles.hConnector} />}
                    </div>

                    <div className={styles.gridCell}>
                        {renderNode(currentEmployee, true)}
                    </div>

                    <div className={styles.gridCell}>
                        {directReports.length > 0 && <div className={styles.hConnector} />}
                    </div>

                    <div className={styles.gridCell}>
                        {directReports.length > 0 && (
                            <div className={styles.reportsStack}>
                                {directReports.map(report => (
                                    <div key={report.id} className={styles.stackItem}>
                                        {renderNode(report, false)}
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
            </div>
        );
    }

    // Default Vertical Tree Layout
    return (
        <div className={styles.orgChartContainer}>
            <Tree
                lineWidth={'2px'}
                lineColor={'#e0e0e0'}
                lineBorderRadius={'10px'}
                label={
                    manager
                        ? renderNode(manager, false, 'Manager')
                        : renderNode(currentEmployee, true)
                }
            >
                {manager ? (
                    <>
                        {buildOrgTree(currentEmployee)}
                        {peers.map(peer => (
                            <TreeNode key={peer.id} label={renderNode(peer, false, 'Peer')} />
                        ))}
                    </>
                ) : (
                    directReports.map(report => buildOrgTree(report))
                )}
            </Tree>
        </div>
    );
};
