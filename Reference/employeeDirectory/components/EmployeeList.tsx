import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployee } from '../models/IEmployee';
import { EmployeeGrid } from './EmployeeGrid';
import { GraphService } from '../services/GraphService';

export interface IEmployeeListProps {
    employees: IEmployee[];
    onUserSelect: (employee: IEmployee) => void;
    graphService: GraphService;
    presence?: { [key: string]: string };
    kudosCounts: { [email: string]: number };
    onGiveKudos: (user: IEmployee) => void;
    onViewKudos: (user: IEmployee) => void;
    isKudosEnabled?: boolean;
    // isDarkTheme?: boolean; // Removed unused prop
    visualTheme?: 'modern-scrolling-premium' | 'modern-tabbed-premium';
}

export const EmployeeList: React.FunctionComponent<IEmployeeListProps> = (props) => {
    return (
        <div className={styles.employeeSection}>
            <div className={styles.container}>
                <EmployeeGrid
                    employees={props.employees}
                    onUserSelect={props.onUserSelect}
                    graphService={props.graphService}
                    presence={props.presence}
                    kudosCounts={props.kudosCounts}
                    onGiveKudos={props.onGiveKudos}
                    onViewKudos={props.onViewKudos}
                    isKudosEnabled={props.isKudosEnabled}
                    visualTheme={props.visualTheme}
                    isListView={true}
                />
            </div>
        </div>
    );
};
