import * as React from 'react';
import {
    TextField,
    PrimaryButton,
    DefaultButton,
    Stack,
    Label,
    ITag,
    TagPicker,
    IBasePickerSuggestionsProps,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType
} from '@fluentui/react';
import { IEmployee } from '../models/IEmployee';
import { GraphService } from '../services/GraphService';
import styles from './EmployeeDirectory.module.scss';

export interface IProfileEditFormProps {
    employee: IEmployee;
    graphService: GraphService;
    editFields?: string[];
    onSave: (updatedEmployee: Partial<IEmployee>) => void;
    onCancel: () => void;
}

export interface IProfileEditFormState {
    aboutMe: string;
    skills: string[];
    interests: string[];
    pastProjects: string[];
    mobilePhone: string;
    businessPhone: string;
    officeLocation: string;
    isLoading: boolean;
    error?: string;
    originalDetails?: Partial<IEmployee>;
}

const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested tags',
    noResultsFoundText: 'No tags found',
};

export class ProfileEditForm extends React.Component<IProfileEditFormProps, IProfileEditFormState> {
    constructor(props: IProfileEditFormProps) {
        super(props);
        this.state = {
            aboutMe: props.employee.aboutMe || '',
            skills: props.employee.skills || [],
            interests: props.employee.interests || [],
            pastProjects: props.employee.pastProjects || [],
            mobilePhone: props.employee.mobilePhone || '',
            businessPhone: (props.employee.businessPhones && props.employee.businessPhones.length > 0) ? props.employee.businessPhones[0] : '',
            officeLocation: props.employee.officeLocation || '',
            isLoading: true
        };
    }

    public componentDidMount(): void {
        this.props.graphService.getUserDetails(this.props.employee.id).then(details => {
            if (details) {
                this.setState((prevState) => ({
                    aboutMe: details.aboutMe || prevState.aboutMe,
                    skills: details.skills || prevState.skills,
                    interests: details.interests || prevState.interests,
                    pastProjects: details.pastProjects || prevState.pastProjects,
                    mobilePhone: details.mobilePhone || prevState.mobilePhone,
                    businessPhone: (details.businessPhones && details.businessPhones.length > 0) ? details.businessPhones[0] : prevState.businessPhone,
                    officeLocation: details.officeLocation || prevState.officeLocation,
                    isLoading: false,
                    originalDetails: details
                }));
            } else {
                this.setState({ isLoading: false });
            }
        }).catch(error => {
            console.error("Error pre-fetching profile info", error);
            this.setState({ isLoading: false, error: "Could not load current profile details. You can still make updates." });
        });
    }

    private readonly _onFilterChanged = (filterText: string, tagList: ITag[] | undefined): ITag[] => {
        return filterText
            ? [
                { key: filterText, name: filterText },
                ... (tagList || []).filter(tag => tag.name.toLowerCase().includes(filterText.toLowerCase()))
            ].filter((tag, index, self) => self.findIndex(t => t.key === tag.key) === index)
            : [];
    };

    private readonly _onSkillsChange = (items?: ITag[]): void => {
        this.setState(prevState => ({ skills: items ? items.map(i => i.name) : [] }));
    };

    private readonly _onInterestsChange = (items?: ITag[]): void => {
        this.setState(prevState => ({ interests: items ? items.map(i => i.name) : [] }));
    };

    private readonly _onProjectsChange = (items?: ITag[]): void => {
        this.setState(prevState => ({ pastProjects: items ? items.map(i => i.name) : [] }));
    };

    private readonly _getProfileUpdates = (originalDetails: Partial<IEmployee> | undefined): Partial<IEmployee> => {
        const updates: Partial<IEmployee> = {};
        const { editFields } = this.props;

        const isAllowed = (field: string): boolean => !editFields || editFields.includes(field);

        const hasArrayChanged = (current: string[], original: string[] | undefined): boolean => {
            const currStr = [...current].sort((a, b) => a.localeCompare(b)).join(',');
            const origStr = [...(original || [])].sort((a, b) => a.localeCompare(b)).join(',');
            return currStr !== origStr;
        };

        // Simple string fields
        const stringFields: { key: keyof IEmployee, stateKey: keyof IProfileEditFormState }[] = [
            { key: 'aboutMe', stateKey: 'aboutMe' },
            { key: 'mobilePhone', stateKey: 'mobilePhone' },
            { key: 'officeLocation', stateKey: 'officeLocation' }
        ];

        stringFields.forEach(field => {
            if (isAllowed(field.key)) {
                const currentVal = this.state[field.stateKey] as string;
                const originalVal = (originalDetails?.[field.key] as string) || '';
                if (currentVal !== originalVal) {
                    updates[field.key] = currentVal as any;
                }
            }
        });

        // Array fields
        const arrayFields: { key: keyof IEmployee, stateKey: keyof IProfileEditFormState }[] = [
            { key: 'skills', stateKey: 'skills' },
            { key: 'interests', stateKey: 'interests' },
            { key: 'pastProjects', stateKey: 'pastProjects' }
        ];

        arrayFields.forEach(field => {
            if (isAllowed(field.key)) {
                const currentVal = this.state[field.stateKey] as string[];
                const originalVal = originalDetails?.[field.key] as string[] | undefined;
                if (hasArrayChanged(currentVal, originalVal)) {
                    updates[field.key] = currentVal as any;
                }
            }
        });

        // Business Phones (Special case)
        const originalBizPhone = (originalDetails?.businessPhones && originalDetails.businessPhones.length > 0) ? originalDetails.businessPhones[0] : '';
        if (isAllowed('businessPhones') && this.state.businessPhone !== originalBizPhone) {
            updates.businessPhones = [this.state.businessPhone];
        }

        return updates;
    };

    private readonly _handleSave = (): void => {
        const updates = this._getProfileUpdates(this.state.originalDetails);

        if (Object.keys(updates).length === 0) {
            this.props.onCancel();
            return;
        }

        this.props.onSave(updates);
    };

    public render(): React.ReactElement<IProfileEditFormProps> {
        const { aboutMe, skills, interests, pastProjects } = this.state;

        const skillTags: ITag[] = skills.map(s => ({ key: s, name: s }));
        const interestTags: ITag[] = interests.map(i => ({ key: i, name: i }));
        const projectTags: ITag[] = pastProjects.map(p => ({ key: p, name: p }));

        const renderPhones = (): JSX.Element | null => {
            const canEditMobile = !this.props.editFields || this.props.editFields.includes('mobilePhone');
            const canEditBusiness = !this.props.editFields || this.props.editFields.includes('businessPhones');

            if (canEditMobile || canEditBusiness) {
                return (
                    <Stack horizontal tokens={{ childrenGap: 20 }}>
                        <Stack.Item grow>
                            <TextField
                                label="Mobile Phone"
                                value={this.state.mobilePhone}
                                onChange={(_e, newValue) => this.setState({ mobilePhone: newValue || '' })}
                                disabled={!canEditMobile}
                                placeholder="+1 (555) 000-0000"
                            />
                        </Stack.Item>
                        <Stack.Item grow>
                            <TextField
                                label="Business Phone"
                                value={this.state.businessPhone}
                                onChange={(_e, newValue) => this.setState({ businessPhone: newValue || '' })}
                                disabled={!canEditBusiness}
                                placeholder="+1 (555) 000-0000"
                            />
                        </Stack.Item>
                    </Stack>
                );
            }
            return null;
        };

        const renderOffice = (): JSX.Element | null => {
            if (!this.props.editFields || this.props.editFields.includes('officeLocation')) {
                return (
                    <TextField
                        label="Office Location"
                        value={this.state.officeLocation}
                        onChange={(_e, newValue) => this.setState({ officeLocation: newValue || '' })}
                        placeholder="e.g. Building 4, Floor 2"
                    />
                );
            }
            return null;
        };

        const renderBio = (): JSX.Element | null => {
            if (!this.props.editFields || this.props.editFields.includes('aboutMe')) {
                return (
                    <div>
                        <Label>Biography</Label>
                        <TextField
                            multiline
                            rows={4}
                            value={aboutMe}
                            onChange={(_e, newValue) => this.setState({ aboutMe: newValue || '' })}
                            placeholder="Tell us about yourself..."
                        />
                    </div>
                );
            }
            return null;
        };

        const renderSkills = (): JSX.Element | null => {
            if (!this.props.editFields || this.props.editFields.includes('skills')) {
                return (
                    <div>
                        <Label>Skills</Label>
                        <TagPicker
                            onResolveSuggestions={this._onFilterChanged}
                            getTextFromItem={(item: ITag) => item.name}
                            pickerSuggestionsProps={pickerSuggestionsProps}
                            selectedItems={skillTags}
                            onChange={this._onSkillsChange}
                            inputProps={{
                                'aria-label': 'Skills picker',
                                placeholder: 'Add skills...'
                            }}
                        />
                    </div>
                );
            }
            return null;
        };

        const renderInterests = (): JSX.Element | null => {
            if (!this.props.editFields || this.props.editFields.includes('interests')) {
                return (
                    <div>
                        <Label>Interests</Label>
                        <TagPicker
                            onResolveSuggestions={this._onFilterChanged}
                            getTextFromItem={(item: ITag) => item.name}
                            pickerSuggestionsProps={pickerSuggestionsProps}
                            selectedItems={interestTags}
                            onChange={this._onInterestsChange}
                            inputProps={{
                                'aria-label': 'Interests picker',
                                placeholder: 'Add interests...'
                            }}
                        />
                    </div>
                );
            }
            return null;
        };

        const renderProjects = (): JSX.Element | null => {
            if (!this.props.editFields || this.props.editFields.includes('pastProjects')) {
                return (
                    <div>
                        <Label>Projects</Label>
                        <TagPicker
                            onResolveSuggestions={this._onFilterChanged}
                            getTextFromItem={(item: ITag) => item.name}
                            pickerSuggestionsProps={pickerSuggestionsProps}
                            selectedItems={projectTags}
                            onChange={this._onProjectsChange}
                            inputProps={{
                                'aria-label': 'Projects picker',
                                placeholder: 'Add projects...'
                            }}
                        />
                    </div>
                );
            }
            return null;
        };

        if (this.state.isLoading) {
            return (
                <div style={{ padding: '40px 0', textAlign: 'center' }}>
                    <Spinner size={SpinnerSize.large} label="Fetching current profile info..." />
                </div>
            );
        }

        return (
            <div className={styles.profileEditForm}>
                <Stack tokens={{ childrenGap: 20 }}>
                    {this.state.error && (
                        <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => this.setState({ error: undefined })}>
                            {this.state.error}
                        </MessageBar>
                    )}
                    {renderPhones()}
                    {renderOffice()}
                    {renderBio()}
                    {renderSkills()}
                    {renderInterests()}
                    {renderProjects()}

                    <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                        <PrimaryButton
                            text="Save Changes"
                            onClick={this._handleSave}
                            iconProps={{ iconName: 'Save' }}
                        />
                        <DefaultButton
                            text="Cancel"
                            onClick={this.props.onCancel}
                        />
                    </Stack>
                </Stack>
            </div>
        );
    }
}
