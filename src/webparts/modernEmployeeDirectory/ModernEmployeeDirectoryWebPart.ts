import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';

import * as strings from 'ModernEmployeeDirectoryWebPartStrings';
import ModernEmployeeDirectory from './components/ModernEmployeeDirectory';
import { IModernEmployeeDirectoryProps } from './components/IModernEmployeeDirectoryProps';
import { GraphService } from './services/GraphService';

export interface IModernEmployeeDirectoryWebPartProps {
  description: string;
  profileLayout: 'scroll' | 'tab';
  mainHeadingSize: number;
  subHeadingSize: number;
  contentHeadingSize: number;
  // Kudos System
  enableKudos: boolean;
  kudosListId: string;
  kudosRecipientColumn: string;
  kudosAuthorColumn: string;
  kudosMessageColumn: string;
  kudosBadgeTypeColumn: string;
  orgChartLayout: 'vertical' | 'horizontal' | 'compact';
  updatableFields: string[];
  enablePagination: boolean;
  pageSize: number;
  paginationType: 'loadMore' | 'prevNext';
  // Hall of Fame
  featuredPeople: IPropertyFieldGroupOrPerson[];
  minKudosCount: number;
  filterType: 'none' | 'domain' | 'extension' | 'department' | 'location';
  filterValue: string;
  filterSecondaryValue: string;
  homePageFilterFields: string[];
  // Audit Logging
  enableAudit: boolean;
  auditListId: string;
  auditActivityColumn: string;
  auditActorColumn: string;
  auditTargetColumn: string;
  auditDetailsColumn: string;
  enableAuditDebug: boolean;
  containerMargin: number;
  badgeCircleSize: number;
  badgeFontSize: number;

}

export default class ModernEmployeeDirectoryWebPart extends BaseClientSideWebPart<IModernEmployeeDirectoryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _graphService: GraphService | undefined;
  private _departments: string[] = [];
  private _locations: string[] = [];
  private _loadingFilters: boolean = false;

  public render(): void {
    const element: React.ReactElement<IModernEmployeeDirectoryProps> = React.createElement(
      ModernEmployeeDirectory,
      {
        ...this.properties,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        featuredEmails: this.properties.featuredPeople ? this.properties.featuredPeople.map(p => p.email).filter(e => !!e) as string[] : [],

        // Defaults for values that might be missing in properties
        profileLayout: this.properties.profileLayout || 'scroll',
        mainHeadingSize: this.properties.mainHeadingSize || 28,
        subHeadingSize: this.properties.subHeadingSize || 24,
        contentHeadingSize: this.properties.contentHeadingSize || 14,
        enableKudos: this.properties.enableKudos || false,
        kudosRecipientColumn: this.properties.kudosRecipientColumn || 'Recipient',
        kudosAuthorColumn: this.properties.kudosAuthorColumn || 'Author',
        kudosMessageColumn: this.properties.kudosMessageColumn || 'Message',
        kudosBadgeTypeColumn: this.properties.kudosBadgeTypeColumn || 'BadgeType',
        orgChartLayout: this.properties.orgChartLayout || 'vertical',
        updatableFields: this.properties.updatableFields || [],
        enablePagination: this.properties.enablePagination || false,
        pageSize: this.properties.pageSize || 10,
        paginationType: this.properties.paginationType || 'loadMore',
        minKudosCount: this.properties.minKudosCount || 0,
        filterType: this.properties.filterType || 'none',
        filterValue: this.properties.filterValue || '',
        filterSecondaryValue: this.properties.filterSecondaryValue || '',
        homePageFilterFields: this.properties.homePageFilterFields || [],
        enableAudit: this.properties.enableAudit || false,
        auditActivityColumn: this.properties.auditActivityColumn || 'Activity',
        auditActorColumn: this.properties.auditActorColumn || 'Actor',
        auditTargetColumn: this.properties.auditTargetColumn || 'Target',
        auditDetailsColumn: this.properties.auditDetailsColumn || 'Details',
        enableAuditDebug: this.properties.enableAuditDebug || false,
        containerMargin: Number(this.properties.containerMargin) || 0,
        badgeCircleSize: this.properties.badgeCircleSize || 32,
        badgeFontSize: this.properties.badgeFontSize || 12,
      }
    );

    ReactDom.render(element, this.domElement);
  }



  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    // Force re-render when properties change
    // This ensures immediate UI updates for layout and data source changes
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected onInit(): Promise<void> {
    this._graphService = new GraphService(this.context);
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._departments.length > 0 || this._locations.length > 0 || this._loadingFilters) {
      return;
    }

    this._fetchUniqueFilterValues().catch(error => { console.error(error); });
  }

  private async _fetchUniqueFilterValues(): Promise<void> {
    if (!this._graphService) return;

    this._loadingFilters = true;
    this.context.propertyPane.refresh();

    try {
      const [departments, locations] = await Promise.all([
        this._graphService.getUniqueValues('department'),
        this._graphService.getUniqueValues('officeLocation')
      ]);

      this._departments = departments;
      this._locations = locations;
    } catch (error) {
      console.error('[WebPart] Error fetching unique filter values:', error);
    } finally {
      this._loadingFilters = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _resolveColumnInternalName(listId: string, fieldId: string): Promise<string> {
    if (!listId || !fieldId) return '';
    try {
      // Import PnP dynamically
      const { spfi } = await import(/* webpackChunkName: 'pnp-sp' */ '@pnp/sp');
      const { SPFx } = await import(/* webpackChunkName: 'pnp-sp' */ '@pnp/sp');
      const sp = spfi().using(SPFx(this.context));

      const field = await sp.web.lists.getById(listId).fields.getById(fieldId).select('InternalName')();
      return field ? field.InternalName : '';
    } catch (error) {
      console.error('[WebPart] Error resolving column internal name:', error);
      return '';
    }
  }

  private async _onColumnChange(propertyPath: string, oldValue: unknown, newValue: unknown, listId: string): Promise<void> {
    if (!newValue) return;

    // If a column GUID is selected, resolve it to internal name
    if (listId) {
      const internalName = await this._resolveColumnInternalName(listId, newValue as string);
      // Save internal name if resolved, otherwise keep original value (GUID)
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, internalName || (newValue as string));
    } else {
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
    this.context.propertyPane.refresh();
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
      this.domElement.style.setProperty('--primary-color', currentTheme.palette?.themePrimary || '#0078d4');
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "General configuration for the directory display and basic behavior."
          },
          groups: [
            {
              groupName: "Display Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('containerMargin', {
                  label: 'Container Margin (px)',
                  options: [
                    { key: 0, text: '0' },
                    { key: 2, text: '2' },
                    { key: 4, text: '4' },
                    { key: 6, text: '6' },
                    { key: 8, text: '8' },
                    { key: 10, text: '10' },
                    { key: 12, text: '12' },
                    { key: 14, text: '14' },
                    { key: 16, text: '16' },
                    { key: 18, text: '18' },
                    { key: 20, text: '20' },
                    { key: 22, text: '22' },
                    { key: 24, text: '24' },
                    { key: 26, text: '26' },
                    { key: 28, text: '28' },
                    { key: 30, text: '30' }
                  ],
                  selectedKey: this.properties.containerMargin || 0
                }),
                PropertyPaneSlider('badgeCircleSize', {
                  label: 'Badge Circle Size',
                  min: 20,
                  max: 60,
                  step: 1,
                  value: this.properties.badgeCircleSize || 32,
                  showValue: true
                }),
                PropertyPaneSlider('badgeFontSize', {
                  label: 'Badge Font Size',
                  min: 8,
                  max: 20,
                  step: 1,
                  value: this.properties.badgeFontSize || 12,
                  showValue: true
                }),

              ]
            },
            {
              groupName: "Layout & Theme",
              groupFields: [
                PropertyPaneDropdown('profileLayout', {
                  label: 'Profile Layout',
                  options: [
                    { key: 'scroll', text: 'Scrolling' },
                    { key: 'tab', text: 'Tabbed' }
                  ]
                }),
                PropertyPaneDropdown('orgChartLayout', {
                  label: 'Org Chart Layout',
                  options: [
                    { key: 'vertical', text: 'Vertical Tree' },
                    { key: 'horizontal', text: 'Horizontal Tree' },
                    { key: 'compact', text: 'Compact List' }
                  ]
                })
              ]
            },
            {
              groupName: 'Font Size Settings',
              groupFields: [
                PropertyPaneSlider('mainHeadingSize', {
                  label: 'Homepage Title Font Size',
                  min: 20,
                  max: 40,
                  step: 1,
                  value: this.properties.mainHeadingSize || 28,
                  showValue: true
                }),
                PropertyPaneSlider('subHeadingSize', {
                  label: 'Detail Page Title Font Size (Profile Page)',
                  min: 16,
                  max: 32,
                  step: 1,
                  value: this.properties.subHeadingSize || 24,
                  showValue: true
                }),
                PropertyPaneSlider('contentHeadingSize', {
                  label: 'Section Heading Font Size (User Name)',
                  min: 12,
                  max: 24,
                  step: 1,
                  value: this.properties.contentHeadingSize || 14,
                  showValue: true
                })
              ]
            },
            {
              groupName: "Navigation & Pagination",
              groupFields: [
                PropertyPaneToggle('enablePagination', {
                  label: 'Enable Pagination',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneSlider('pageSize', {
                  label: 'Users per page',
                  min: 5,
                  max: 50,
                  step: 1,
                  value: this.properties.pageSize || 10,
                  showValue: true,
                  disabled: !this.properties.enablePagination
                }),
                PropertyPaneDropdown('paginationType', {
                  label: 'Pagination style',
                  options: [
                    { key: 'loadMore', text: 'Load More' },
                    { key: 'prevNext', text: 'Previous / Next' }
                  ],
                  selectedKey: this.properties.paginationType || 'loadMore',
                  disabled: !this.properties.enablePagination
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Configure core directory features including filters, recognition, and the Hall of Fame."
          },
          groups: [
            {
              groupName: 'Organization Filters',
              groupFields: [
                PropertyPaneDropdown('filterType', {
                  label: 'Filter Type',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'department', text: 'By Department' },
                    { key: 'location', text: 'By Office Location' },
                    { key: 'domain', text: 'By Email Domain' },
                    { key: 'extension', text: 'By Extension Attribute' }
                  ],
                  selectedKey: this.properties.filterType || 'none'
                }),
                ...(this.properties.filterType && this.properties.filterType !== 'none' ? [
                  (this.properties.filterType === 'department' || this.properties.filterType === 'location') ?
                    PropertyPaneDropdown('filterValue', {
                      label: this._getFilterValueLabel(),
                      options: (this.properties.filterType === 'department' ? this._departments : this._locations)
                        .map(v => ({ key: v, text: v })),
                      disabled: this._loadingFilters
                    }) :
                    PropertyPaneTextField('filterValue', {
                      label: this._getFilterValueLabel()
                    })
                ] : []),
                ...(this.properties.filterType === 'extension' ? [
                  PropertyPaneTextField('filterSecondaryValue', {
                    label: 'Attribute Value'
                  })
                ] : []),
                PropertyFieldMultiSelect('homePageFilterFields', {
                  key: 'homePageFilterFields',
                  label: 'Home Page Dropdown Filters',
                  options: [
                    { key: 'department', text: 'Department' },
                    { key: 'officeLocation', text: 'Office Location' },
                    { key: 'jobTitle', text: 'Job Title' },
                    { key: 'city', text: 'City' },
                    { key: 'state', text: 'State/Province' },
                    { key: 'country', text: 'Country' }
                  ],
                  selectedKeys: this.properties.homePageFilterFields || []
                })
              ]
            },
            {
              groupName: 'Kudos System',
              groupFields: [
                PropertyPaneToggle('enableKudos', {
                  label: 'Enable Kudos',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneSlider('minKudosCount', {
                  label: 'Min Kudos for Hall of Fame',
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.properties.minKudosCount || 0,
                  showValue: true,
                  disabled: !this.properties.enableKudos
                }),
                ...(this.properties.enableKudos ? [
                  PropertyFieldListPicker('kudosListId', {
                    label: 'Select Kudos List',
                    selectedList: this.properties.kudosListId,
                    includeHidden: false,
                    orderBy: PropertyFieldListPickerOrderBy.Title,
                    disabled: false,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    context: this.context as any,
                    deferredValidationTime: 0,
                    key: 'kudosListPickerFieldId'
                  }),
                  ...(this.properties.kudosListId ? [
                    PropertyFieldColumnPicker('kudosRecipientColumn', {
                      label: 'Recipient Column',
                      context: this.context as unknown as any,
                      selectedColumn: this.properties.kudosRecipientColumn,
                      listId: this.properties.kudosListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.kudosListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'kudosRecipientColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyFieldColumnPicker('kudosAuthorColumn', {
                      label: 'Author Column',
                      context: this.context as unknown as any,
                      selectedColumn: this.properties.kudosAuthorColumn,
                      listId: this.properties.kudosListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.kudosListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'kudosAuthorColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyFieldColumnPicker('kudosMessageColumn', {
                      label: 'Message Column',
                      context: this.context as unknown as any,
                      selectedColumn: this.properties.kudosMessageColumn,
                      listId: this.properties.kudosListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.kudosListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'kudosMessageColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyFieldColumnPicker('kudosBadgeTypeColumn', {
                      label: 'Badge Type Column',
                      context: this.context as unknown as any,
                      selectedColumn: this.properties.kudosBadgeTypeColumn,
                      listId: this.properties.kudosListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.kudosListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'kudosBadgeTypeColumnPickerFieldId',
                      displayHiddenColumns: false
                    })
                  ] : [])
                ] : [])
              ]
            },
            {
              groupName: 'Hall of Fame',
              groupFields: [
                PropertyFieldPeoplePicker('featuredPeople', {
                  label: 'Manually Featured People',
                  initialData: this.properties.featuredPeople,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null as any,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Advanced management settings including self-service profile updates and audit logging."
          },
          groups: [
            {
              groupName: "Self-Service Settings",
              groupFields: [
                PropertyFieldMultiSelect('updatableFields', {
                  key: 'updatableFields',
                  label: 'Updatable Profile Fields',
                  options: [
                    { key: 'jobTitle', text: 'Job Title' },
                    { key: 'aboutMe', text: 'Bio (About Me)' },
                    { key: 'mobilePhone', text: 'Mobile Phone' },
                    { key: 'officeLocation', text: 'Office Location' },
                    { key: 'skills', text: 'Skills' },
                    { key: 'interests', text: 'Interests' },
                    { key: 'pastProjects', text: 'Past Projects' }
                  ],
                  selectedKeys: this.properties.updatableFields
                })
              ]
            },
            {
              groupName: 'Audit & Insights',
              groupFields: [
                PropertyPaneToggle('enableAudit', {
                  label: 'Enable Audit Logging',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                ...(this.properties.enableAudit ? [
                  PropertyFieldListPicker('auditListId', {
                    label: 'Select Audit List',
                    selectedList: this.properties.auditListId,
                    includeHidden: false,
                    orderBy: PropertyFieldListPickerOrderBy.Title,
                    disabled: false,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    context: this.context as any,
                    deferredValidationTime: 0,
                    key: 'auditListPickerFieldId'
                  }),
                  ...(this.properties.auditListId ? [
                    PropertyFieldColumnPicker('auditActivityColumn', {
                      label: 'Activity Column',
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      context: this.context as any,
                      selectedColumn: this.properties.auditActivityColumn,
                      listId: this.properties.auditListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.auditListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'auditActivityColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyFieldColumnPicker('auditActorColumn', {
                      label: 'Actor (Person) Column',
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      context: this.context as any,
                      selectedColumn: this.properties.auditActorColumn,
                      listId: this.properties.auditListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.auditListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'auditActorColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyFieldColumnPicker('auditTargetColumn', {
                      label: 'Target (Text/Person) Column',
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      context: this.context as any,
                      selectedColumn: this.properties.auditTargetColumn,
                      listId: this.properties.auditListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.auditListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'auditTargetColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyFieldColumnPicker('auditDetailsColumn', {
                      label: 'Details (Multi-line) Column',
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      context: this.context as any,
                      selectedColumn: this.properties.auditDetailsColumn,
                      listId: this.properties.auditListId,
                      disabled: false,
                      orderBy: PropertyFieldColumnPickerOrderBy.Title,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      onPropertyChange: (p: string, o: unknown, n: unknown) => { this._onColumnChange(p, o, n, this.properties.auditListId).catch((err: Error) => { console.error(err); }); },
                      properties: this.properties,
                      deferredValidationTime: 0,
                      key: 'auditDetailsColumnPickerFieldId',
                      displayHiddenColumns: false
                    }),
                    PropertyPaneToggle('enableAuditDebug', {
                      label: 'Show Audit Debug Panel',
                      onText: 'Show',
                      offText: 'Hide'
                    })
                  ] : [])
                ] : [])
              ]
            }
          ]
        }
      ]
    };
  }

  private _getFilterValueLabel(): string {
    switch (this.properties.filterType) {
      case 'department':
        return 'Department Name';
      case 'location':
        return 'Office Location';
      case 'domain':
        return 'Domain (e.g. contoso.com)';
      case 'extension':
        return 'Attribute Name (e.g. CustomAttribute1)';
      default:
        return 'Filter Value';
    }
  }

}
