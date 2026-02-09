import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';

export interface IFiltersProps {
    onSearch: (searchTerm: string) => void;
    onFilter: (filters: { [key: string]: string }) => void;
    filterFields: string[];
    filterOptions: { [key: string]: string[] };
    searchQuery?: string;
}

export const Filters: React.FunctionComponent<IFiltersProps> = (props) => {
    const [currentFilters, setCurrentFilters] = React.useState<{ [key: string]: string }>({});

    const _handleFilterChange = (field: string, value: string) => {
        const newFilters = { ...currentFilters, [field]: value };
        setCurrentFilters(newFilters);
        props.onFilter(newFilters);
    };

    return (
        <div className={styles.searchBar}>
            <input
                type="text"
                className={styles.searchInput}
                placeholder="Search employees by name, title, or department..."
                value={props.searchQuery || ''}
                onChange={(e) => props.onSearch(e.target.value)}
            />

            {props.filterFields.map(field => {
                const options = props.filterOptions[field] || [];
                let label = field.charAt(0).toUpperCase() + field.slice(1);
                if (field === 'officeLocation') label = 'Location';

                return (
                    <select
                        key={field}
                        className={styles.filterDropdown}
                        onChange={(e) => _handleFilterChange(field, e.target.value)}
                    >
                        <option value="">All {label}s</option>
                        {options.map(opt => (
                            <option key={opt} value={opt}>{opt}</option>
                        ))}
                    </select>
                );
            })}
        </div>
    );
};
