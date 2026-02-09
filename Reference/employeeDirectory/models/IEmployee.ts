export interface IEmployee {
  id: string;
  displayName: string;
  jobTitle: string;
  mail: string;
  mobilePhone: string;
  businessPhones: string[];
  officeLocation: string;
  department: string;
  companyName?: string;
  city?: string;
  state?: string;
  country?: string;
  userPrincipalName?: string;
  photoUrl?: string;
  manager?: IEmployee;
  directReports?: IEmployee[];
  onPremisesExtensionAttributes?: {
    [key: string]: string | undefined;
  };
  // Expanded M365 Profile Properties
  aboutMe?: string;
  birthday?: string;
  hireDate?: string;
  skills?: string[];
  interests?: string[];
  responsibilities?: string[];
  schools?: string[];
  pastProjects?: string[];
  businessAddress?: {
    street?: string;
    city?: string;
    state?: string;
    country?: string;
    postalCode?: string;
  };
}
