import { IEmployee } from '../components/Home/DirectoryHome';

export interface ICurrentUser {
    initials: string;
    name: string;
    email: string;
}

export class MockDataService {
    private static readonly employees: IEmployee[] = [
        {
            id: 'AlexW@devtenant0424.onmicrosoft.com',
            displayName: 'Sarah Johnson',
            givenName: 'Sarah',
            surname: 'Johnson',
            jobTitle: 'Chief Executive Officer',
            department: 'Executive',
            mail: 'AlexW@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'AlexW@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 123-4567'],
            mobilePhone: '+1 (555) 123-4568',
            officeLocation: 'Building 1, Floor 5',
            city: 'Seattle',
            state: 'WA',
            country: 'United States',
            initials: 'SJ',
            presence: { availability: 'Available', activity: 'Available' },
            isOnline: true,
            colleagues: ['a2c4e6f8-1b3d-5e7a-9c0b-2d4f6a8c0e2b', 'f3d5b7c9-2e4a-6c8d-0f1b-3e5a7c9d1f3b']
            // CEO - no manager
        },
        {
            id: 'anu@vishpowerlabs.com',
            displayName: 'Michael Chen',
            givenName: 'Michael',
            surname: 'Chen',
            jobTitle: 'VP of Engineering',
            department: 'Engineering',
            mail: 'anu@vishpowerlabs.com',
            userPrincipalName: 'anu@vishpowerlabs.com',
            businessPhones: ['+1 (555) 234-5678'],
            mobilePhone: '+1 (555) 234-5679',
            officeLocation: 'Building 2, Floor 3',
            city: 'San Francisco',
            state: 'CA',
            country: 'United States',
            initials: 'MC',
            presence: { availability: 'Busy', activity: 'InAMeeting' },
            isOnline: false,
            managerId: 'e7b3a8d1-4c2f-4a5e-9b1c-3d8e7f2a1b4c', // Reports to CEO
            colleagues: ['f3d5b7c9-2e4a-6c8d-0f1b-3e5a7c9d1f3b', 'b4d6f8a0-3c5e-7a9b-1d3f-5e7a9b1d3f5a']
        },
        {
            id: 'apiuser@devtenant0424.onmicrosoft.com',
            displayName: 'Emily Rodriguez',
            givenName: 'Emily',
            surname: 'Rodriguez',
            jobTitle: 'VP of Design',
            department: 'Design',
            mail: 'apiuser@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'apiuser@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 345-6789'],
            mobilePhone: '+1 (555) 345-6790',
            officeLocation: 'Building 3, Floor 2',
            city: 'Austin',
            state: 'TX',
            country: 'United States',
            initials: 'ER',
            presence: { availability: 'Available', activity: 'Available' },
            isOnline: true,
            managerId: 'e7b3a8d1-4c2f-4a5e-9b1c-3d8e7f2a1b4c' // Reports to CEO
        },
        {
            id: 'Blossom@devtenant0424.onmicrosoft.com',
            displayName: 'David Kim',
            givenName: 'David',
            surname: 'Kim',
            jobTitle: 'Senior Product Manager',
            department: 'Product',
            mail: 'Blossom@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'Blossom@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 456-7890'],
            mobilePhone: '+1 (555) 456-7891',
            officeLocation: 'Building 2, Floor 3',
            city: 'New York',
            state: 'NY',
            country: 'United States',
            initials: 'DK',
            presence: { availability: 'Away', activity: 'Away' },
            isOnline: false,
            managerId: 'a2c4e6f8-1b3d-5e7a-9c0b-2d4f6a8c0e2b' // Reports to VP of Engineering
        },
        {
            id: 'demo@vishpowerlabs.com',
            displayName: 'Jessica Martinez',
            givenName: 'Jessica',
            surname: 'Martinez',
            jobTitle: 'Lead Software Engineer',
            department: 'Engineering',
            mail: 'demo@vishpowerlabs.com',
            userPrincipalName: 'demo@vishpowerlabs.com',
            businessPhones: ['+1 (555) 567-8901'],
            mobilePhone: '+1 (555) 567-8902',
            officeLocation: 'Building 4, Floor 1',
            city: 'Chicago',
            state: 'IL',
            country: 'United States',
            initials: 'JM',
            presence: { availability: 'Available', activity: 'Available' },
            isOnline: true,
            managerId: 'e7b3a8d1-4c2f-4a5e-9b1c-3d8e7f2a1b4c' // Reports to CEO
        },
        {
            id: 'demo2@devtenant0424.onmicrosoft.com',
            displayName: 'Robert Taylor',
            givenName: 'Robert',
            surname: 'Taylor',
            jobTitle: 'Senior Data Scientist',
            department: 'Analytics',
            mail: 'demo2@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'demo2@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 678-9012'],
            mobilePhone: '+1 (555) 678-9013',
            officeLocation: 'Building 2, Floor 4',
            city: 'Boston',
            state: 'MA',
            country: 'United States',
            initials: 'RT',
            presence: { availability: 'Offline', activity: 'Offline' },
            isOnline: false,
            managerId: 'a2c4e6f8-1b3d-5e7a-9c0b-2d4f6a8c0e2b' // Reports to VP of Engineering
        },
        {
            id: 'HenriettaM@devtenant0424.onmicrosoft.com',
            displayName: 'Amanda White',
            givenName: 'Amanda',
            surname: 'White',
            jobTitle: 'VP of Human Resources',
            department: 'Human Resources',
            mail: 'HenriettaM@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'HenriettaM@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 789-0123'],
            mobilePhone: '+1 (555) 789-0124',
            officeLocation: 'Building 1, Floor 3',
            city: 'Denver',
            state: 'CO',
            country: 'United States',
            initials: 'AW',
            presence: { availability: 'Available', activity: 'Available' },
            isOnline: true,
            managerId: 'e7b3a8d1-4c2f-4a5e-9b1c-3d8e7f2a1b4c' // Reports to CEO
        },
        {
            id: 'Info@vishpowerlabs.com',
            displayName: 'James Anderson',
            givenName: 'James',
            surname: 'Anderson',
            jobTitle: 'VP of Sales',
            department: 'Sales',
            mail: 'Info@vishpowerlabs.com',
            userPrincipalName: 'Info@vishpowerlabs.com',
            businessPhones: ['+1 (555) 890-1234'],
            mobilePhone: '+1 (555) 890-1235',
            officeLocation: 'Building 5, Floor 2',
            city: 'Atlanta',
            state: 'GA',
            country: 'United States',
            initials: 'JA',
            presence: { availability: 'Busy', activity: 'InACall' },
            isOnline: false,
            managerId: 'e7b3a8d1-4c2f-4a5e-9b1c-3d8e7f2a1b4c' // Reports to CEO
        },
        {
            id: 'LidiaH@devtenant0424.onmicrosoft.com',
            displayName: 'Lisa Thompson',
            givenName: 'Lisa',
            surname: 'Thompson',
            jobTitle: 'Senior QA Engineer',
            department: 'Engineering',
            mail: 'LidiaH@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'LidiaH@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 901-2345'],
            mobilePhone: '+1 (555) 901-2346',
            officeLocation: 'Building 2, Floor 3',
            city: 'Portland',
            state: 'OR',
            country: 'United States',
            initials: 'LT',
            presence: { availability: 'Available', activity: 'Available' },
            isOnline: true,
            managerId: 'a2c4e6f8-1b3d-5e7a-9c0b-2d4f6a8c0e2b' // Reports to VP of Engineering
        },
        {
            id: 'JohannaL@devtenant0424.onmicrosoft.com',
            displayName: 'Lisa Brown',
            givenName: 'Lisa',
            surname: 'Brown',
            jobTitle: 'Senior UX Designer',
            department: 'Design',
            mail: 'JohannaL@devtenant0424.onmicrosoft.com',
            userPrincipalName: 'JohannaL@devtenant0424.onmicrosoft.com',
            businessPhones: ['+1 (555) 012-3456'],
            mobilePhone: '+1 (555) 012-3457',
            officeLocation: 'Building 3, Floor 2',
            city: 'Miami',
            state: 'FL',
            country: 'United States',
            initials: 'DB',
            presence: { availability: 'DoNotDisturb', activity: 'Presenting' },
            isOnline: false,
            managerId: 'b3d5e7f9-2c4e-6f8a-0d1c-3e5f7a9c1d3e' // Reports to VP of Design
        }
    ];


    private static readonly currentUser: ICurrentUser = {
        initials: 'CU',
        name: 'Current User',
        email: 'current.user@contoso.com'
    };

    public static getEmployees(): IEmployee[] {
        return this.employees;
    }

    public static getCurrentUser(): ICurrentUser {
        return this.currentUser;
    }

    public static getEmployeeById(id: string): IEmployee | undefined {
        return this.employees.find(emp => emp.id === id);
    }
}
