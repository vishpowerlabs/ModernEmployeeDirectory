declare interface IEmployeeDirectoryWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    DescriptionFieldLabel: string;
    DomainFilterLabel: string;
    ExcludeFilterLabel: string;
    AppLocalEnvironmentSharePoint: string;
    AppLocalEnvironmentTeams: string;
    AppLocalEnvironmentOffice: string;
    AppSharePointEnvironment: string;
    AppTeamsTabEnvironment: string;
    AppOfficeEnvironment: string;
    AppOutlookEnvironment: string;
    FilterFieldsLabel: string;
}

declare module 'EmployeeDirectoryWebPartStrings' {
    const strings: IEmployeeDirectoryWebPartStrings;
    export = strings;
}
