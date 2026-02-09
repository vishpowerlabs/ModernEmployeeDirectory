import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployee } from '../models/IEmployee';
import { EmployeeCard } from './EmployeeCard';
import { GraphService } from '../services/GraphService';

export interface IEmployeeGridProps {
    employees: IEmployee[];
    onUserSelect: (employee: IEmployee) => void;
    graphService: GraphService;
    presence?: { [key: string]: string };
    kudosCounts: { [email: string]: number };
    onGiveKudos: (user: IEmployee) => void;
    onViewKudos: (user: IEmployee) => void;
    isKudosEnabled?: boolean;
    visualTheme?: 'modern-scrolling-premium' | 'modern-tabbed-premium';
    isListView?: boolean;
}

export const EmployeeGrid: React.FunctionComponent<IEmployeeGridProps> = (props) => {
    return (
        <div className={`${styles.employeeGrid} ${props.isListView ? styles.listView : ''}`}>
            {props.employees.map((emp, index) => (
                <EmployeeCard
                    key={emp.id}
                    employee={emp}
                    onClick={props.onUserSelect}
                    graphService={props.graphService}
                    presenceStatus={props.presence ? props.presence[emp.id] : undefined}
                    animationDelay={`${index * 0.05}s`}
                    kudosCounts={props.kudosCounts}
                    onGiveKudos={props.onGiveKudos}
                    onViewKudos={props.onViewKudos}
                    isKudosEnabled={props.isKudosEnabled}
                    visualTheme={props.visualTheme}
                    isListView={props.isListView}
                />
            ))}
        </div>
    );
};
