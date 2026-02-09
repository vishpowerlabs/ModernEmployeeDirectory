import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'EmployeeDirectoryWebPartStrings';
import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps } from './components/IEmployeeDirectoryProps';
import { GraphService } from './services/GraphService';

export interface IEmployeeDirectoryWebPartProps {
  description: string;
  excludeUsers: string;
  filterFields: string[];
  cardFields: string[];
  filterType: 'none' | 'domain' | 'extension' | 'department' | 'location';
  filterValue: string;
  filterSecondaryValue: string;
  title: string;
  enableStarred: boolean;
  starredUsers: IPropertyFieldGroupOrPerson[];
  titleFontSize: number;
  descriptionFontSize: number;
  visualTheme: 'modern-scrolling-premium' | 'modern-tabbed-premium';
  subHeadingFontSize: number;
  contentFontSize: number;
  headerPadding: number;
  profileHeaderPadding: number;
  profileHeaderHeight: number;
  enableSelfUpdate: boolean;
  editFields: string[];
  headerAlignment: 'left' | 'center' | 'right';
  themeTemplate: 'default' | 'berry' | 'orange' | 'grape' | 'forest' | 'ocean' | 'midnight' | 'mint';
  kudosListId: string;
  kudosTargetField: string;
  kudosMessageField: string;
  kudosTypeField: string;
  kudosDateField: string;
  enableKudos: boolean;
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {
  private _hasRendered: boolean = false;
  private _graphService: GraphService | undefined;
  private _availableProperties: string[] = [];
  private _lists: { key: string; text: string }[] = [];
  private _loadingLists: boolean = false;
  private _kudosColumns: { key: string; text: string }[] = [];
  private _loadingColumns: boolean = false;

  public render(): void {
    const element: React.ReactElement<IEmployeeDirectoryProps> = React.createElement(
      EmployeeDirectory,
      {
        description: this.properties.description,
        excludeUsers: this.properties.excludeUsers,
        filterFields: this.properties.filterFields || ['department', 'officeLocation'],
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: Boolean(this.context.sdks?.microsoftTeams),
        userDisplayName: this.context.pageContext.user.displayName,
        cardFields: this.properties.cardFields || ['email', 'phone', 'location'],
        context: this.context,
        filterType: this.properties.filterType,
        filterValue: this.properties.filterValue,
        filterSecondaryValue: this.properties.filterSecondaryValue,
        title: this.properties.title,
        enableStarred: this.properties.enableStarred,
        starredUsers: this.properties.starredUsers,
        titleFontSize: this.properties.titleFontSize,
        descriptionFontSize: this.properties.descriptionFontSize,


        profileHeaderHeight: this.properties.profileHeaderHeight,
        visualTheme: (this.properties.visualTheme === 'modern-tabbed-premium' ? 'modern-tabbed-premium' : 'modern-scrolling-premium'),
        subHeadingFontSize: this.properties.subHeadingFontSize,
        contentFontSize: this.properties.contentFontSize,
        headerPadding: this.properties.headerPadding,
        profileHeaderPadding: this.properties.profileHeaderPadding,
        enableSelfUpdate: this.properties.enableSelfUpdate,
        editFields: this.properties.editFields || ['aboutMe', 'skills', 'interests', 'pastProjects', 'mobilePhone', 'businessPhones', 'officeLocation'],
        headerAlignment: this.properties.headerAlignment || 'center',
        themeTemplate: this.properties.themeTemplate,
        enableKudos: this.properties.enableKudos,
        kudosListId: this.properties.kudosListId,
        kudosTargetField: this.properties.kudosTargetField,
        kudosMessageField: this.properties.kudosMessageField,
        kudosTypeField: this.properties.kudosTypeField,
        kudosDateField: this.properties.kudosDateField
      }
    );

    ReactDom.render(element, this.domElement);
    this._hasRendered = true;
  }

  protected onInit(): Promise<void> {
    this._graphService = new GraphService(this.context);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      return (this._graphService ? this._graphService.getAvailableProperties() : Promise.resolve([])).then(props => {
        this._availableProperties = props;
      });
    });
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks?.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // office.com
              environmentMessage = strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // Outlook
              environmentMessage = strings.AppOutlookEnvironment;
              break;
            case 'Teams': // Teams
            case 'TeamsModern':
              environmentMessage = strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.AppSharePointEnvironment;
          }
          return environmentMessage;
        }).catch(() => {
          return strings.AppSharePointEnvironment;
        });
    }

    return Promise.resolve(strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: any): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = currentTheme.isInverted;
    this._applyThemeStyles(currentTheme);

    if (this._hasRendered) {
      this.render();
    }
  }

  private _applyThemeStyles(theme: any): void {
    if (!theme) {
      return;
    }

    try {
      const {
        semanticColors,
        palette
      } = theme;

      if (semanticColors) {
        this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
        this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || '');
        this.domElement.style.setProperty('--link', semanticColors.link || '');
        this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
      }

      if (palette) {
        this.domElement.style.setProperty('--themePrimary', palette.themePrimary || '');
        this.domElement.style.setProperty('--themeSecondary', palette.themeSecondary || '');
        this.domElement.style.setProperty('--themeTertiary', palette.themeTertiary || '');
        this.domElement.style.setProperty('--themeLight', palette.themeLight || '');
        this.domElement.style.setProperty('--themeDark', palette.themeDark || '');
        this.domElement.style.setProperty('--themeLighterAlt', palette.themeLighterAlt || '');
        this.domElement.style.setProperty('--white', palette.white || '');
        this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || '');
        this.domElement.style.setProperty('--neutralLight', palette.neutralLight || '');
        this.domElement.style.setProperty('--neutralPrimary', palette.neutralPrimary || '');
        this.domElement.style.setProperty('--neutralSecondary', palette.neutralSecondary || '');
        this.domElement.style.setProperty('--neutralTertiary', palette.neutralTertiary || '');
      }
    } catch (err) {
      console.error("Error applying theme styles", err);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'visualTheme') {
      // visualTheme changes valid
    }
    if (propertyPath === 'kudosListId' && newValue) {
      this._fetchKudosColumns(newValue);
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _fetchKudosColumns(listId: string): void {
    if (!listId) return;
    this._loadingColumns = true;
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/Fields?$filter=Hidden eq false and ReadOnlyField eq false`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then(data => {
        this._kudosColumns = (data.value || []).map((field: any) => ({
          key: field.InternalName,
          text: field.Title
        })).sort((a: any, b: any) => a.text.localeCompare(b.text));
        this._loadingColumns = false;
        this.context.propertyPane.refresh();
      })
      .catch(err => {
        console.error("Error fetching columns", err);
        this._loadingColumns = false;
      });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._lists.length > 0) {
      if (this.properties.kudosListId && this._kudosColumns.length === 0) {
        this._fetchKudosColumns(this.properties.kudosListId);
      }
      return;
    }

    this._loadingLists = true;
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 100`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then(data => {
        this._lists = (data.value || []).map((list: any) => ({
          key: list.Id,
          text: list.Title
        }));
        this._loadingLists = false;
        this.context.propertyPane.refresh();
      })
      .catch(err => {
        console.error("Error fetching lists", err);
        this._loadingLists = false;
      });

    if (this.properties.kudosListId) {
      this._fetchKudosColumns(this.properties.kudosListId);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneToggle('enableSelfUpdate', {
                  label: "Enable Self-Service Profile Update"
                }),
                PropertyFieldMultiSelect('editFields', {
                  key: 'editFields',
                  label: "Editable Profile Fields",
                  options: [
                    { key: 'aboutMe', text: 'Biography' },
                    { key: 'skills', text: 'Skills' },
                    { key: 'interests', text: 'Interests' },
                    { key: 'pastProjects', text: 'Past Projects' },
                    { key: 'mobilePhone', text: 'Mobile Phone' },
                    { key: 'businessPhones', text: 'Business Phone' },
                    { key: 'officeLocation', text: 'Office Location' }
                  ],
                  selectedKeys: this.properties.editFields || ['aboutMe', 'skills', 'interests', 'pastProjects', 'mobilePhone', 'businessPhones', 'officeLocation'],
                  disabled: !this.properties.enableSelfUpdate
                }),
                PropertyPaneDropdown('themeTemplate', {
                  label: 'Theme Template',
                  options: [
                    { key: 'default', text: 'SharePoint Theme' },
                    { key: 'berry', text: 'Berry (Deep Red)' },
                    { key: 'orange', text: 'Orange (Zesty)' },
                    { key: 'grape', text: 'Grape (Royal Purple)' },
                    { key: 'forest', text: 'Forest (Nature Green)' },
                    { key: 'ocean', text: 'Ocean (Deep Blue)' },
                    { key: 'midnight', text: 'Midnight (Dark Indigo)' },
                    { key: 'mint', text: 'Mint (Fresh Teal)' }
                  ],
                  selectedKey: this.properties.themeTemplate || 'default'
                }),
                PropertyPaneTextField('title', {
                  label: "App Title"
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('visualTheme', {
                  label: 'Visual Theme',
                  options: [
                    { key: 'modern-scrolling-premium', text: 'Dynamic Unified View' },
                    { key: 'modern-tabbed-premium', text: 'Structured Tabbed View' }
                  ],
                  selectedKey: this.properties.visualTheme || 'modern-scrolling-premium'
                }),
                PropertyPaneDropdown('headerAlignment', {
                  label: 'Header Alignment',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'center', text: 'Center' },
                    { key: 'right', text: 'Right' }
                  ],
                  selectedKey: this.properties.headerAlignment || 'center'
                })
              ]
            },
            {
              groupName: "Filtering Configuration",
              groupFields: [
                PropertyPaneDropdown('filterType', {
                  label: 'Filter Type',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'domain', text: 'Domain Name' },
                    { key: 'department', text: 'Department' },
                    { key: 'location', text: 'Location' },
                    { key: 'extension', text: 'Extension Attribute' }
                  ]
                }),
                ...(this.properties.filterType === 'none' ? [] : [PropertyPaneTextField('filterValue', {
                  label: this._getFilterValueLabel()
                })]),
                ...(this.properties.filterType === 'extension' ? [PropertyPaneTextField('filterSecondaryValue', {
                  label: 'Extension Attribute Value'
                })] : []),
                PropertyPaneTextField('excludeUsers', {
                  label: "Exclude Users Title",
                  description: "Comma-separated list of emails or names to exclude"
                }),
                PropertyFieldMultiSelect('filterFields', {
                  key: 'filterFields',
                  label: strings.FilterFieldsLabel,
                  options: this._availableProperties.map(p => ({
                    key: p,
                    text: p.charAt(0).toUpperCase() + p.slice(1).replaceAll(/([A-Z])/g, ' $1')
                  })),
                  selectedKeys: this.properties.filterFields || ['department', 'officeLocation']
                })
              ]
            },
            {
              groupName: "Display Customization",
              groupFields: [
                PropertyPaneToggle('enableStarred', {
                  label: "Enable Key People / Starred"
                }),
                ...(this.properties.enableStarred ? [PropertyFieldPeoplePicker('starredUsers', {
                  key: 'starredUsers',
                  label: 'Select Key People',
                  initialData: this.properties.starredUsers,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                })] : [])
              ]
            },
            {
              groupName: "Typography",
              groupFields: [
                PropertyPaneSlider('titleFontSize', {
                  label: "Title Font Size",
                  min: 12,
                  max: 48,
                  step: 2,
                  value: this.properties.titleFontSize || 32
                }),
                PropertyPaneSlider('descriptionFontSize', {
                  label: "Description Font Size",
                  min: 10,
                  max: 24,
                  step: 1,
                  value: this.properties.descriptionFontSize || 15
                }),
                PropertyPaneSlider('subHeadingFontSize', {
                  label: "Sub-heading Font Size (Usernames/Bar)",
                  min: 14,
                  max: 32,
                  step: 1,
                  value: this.properties.subHeadingFontSize || 20
                }),
                PropertyPaneSlider('contentFontSize', {
                  label: "Content/Button Font Size",
                  min: 10,
                  max: 20,
                  step: 1,
                  value: this.properties.contentFontSize || 14
                })
              ]
            },
            {
              groupName: "Layout & Spacing",
              groupFields: [
                PropertyPaneSlider('headerPadding', {
                  label: "Main Header Height (Padding)",
                  min: 10,
                  max: 100,
                  step: 5,
                  value: this.properties.headerPadding || 40
                }),
                PropertyPaneSlider('profileHeaderPadding', {
                  label: "Profile Bar Height (Padding)",
                  min: 5,
                  max: 60,
                  step: 5,
                  value: this.properties.profileHeaderPadding || 30
                }),
                PropertyPaneSlider('profileHeaderHeight', {
                  label: "Explicit Profile Bar Height (Min-Height)",
                  min: 10,
                  max: 100,
                  step: 5,
                  value: this.properties.profileHeaderHeight || 60
                })
              ]
            },
            {
              groupName: "Kudos Recognition Settings",
              groupFields: [
                PropertyPaneToggle('enableKudos', {
                  label: 'Enable Kudos Recognition'
                }),
                ...(this.properties.enableKudos ? [
                  PropertyPaneDropdown('kudosListId', {
                    label: 'Kudos SharePoint List',
                    options: this._lists,
                    disabled: this._loadingLists
                  }),
                  ...(this.properties.kudosListId ? [
                    PropertyPaneDropdown('kudosTargetField', {
                      label: 'Target User Field',
                      options: this._kudosColumns,
                      disabled: this._loadingColumns
                    }),
                    PropertyPaneDropdown('kudosMessageField', {
                      label: 'Kudos Message Field',
                      options: this._kudosColumns,
                      disabled: this._loadingColumns
                    }),
                    PropertyPaneDropdown('kudosTypeField', {
                      label: 'Kudos Type (Choice) Field',
                      options: this._kudosColumns,
                      disabled: this._loadingColumns
                    }),
                    PropertyPaneDropdown('kudosDateField', {
                      label: 'Kudos Date Field',
                      options: this._kudosColumns,
                      disabled: this._loadingColumns
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
    const { filterType } = this.properties;
    if (filterType === 'extension') return 'Extension Attribute Name (e.g. extensionAttribute1)';
    if (filterType === 'department') return 'Department Name';
    if (filterType === 'location') return 'Location Name';
    return (strings.DomainFilterLabel || 'Domain Name');
  }
}
