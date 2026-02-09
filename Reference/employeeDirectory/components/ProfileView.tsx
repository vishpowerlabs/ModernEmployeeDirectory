import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployee } from '../models/IEmployee';
import { GraphService } from '../services/GraphService';
import { OrgChart } from './OrgChart';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IProfileViewProps {
    employee: IEmployee;
    graphService: GraphService;
    onUserSelect: (employee: IEmployee) => void;
    onBackToHome: () => void;
    visualTheme?: 'modern-scrolling-premium' | 'modern-tabbed-premium';
    kudosCounts: { [email: string]: number };
    onGiveKudos: (user: IEmployee) => void;
    onViewKudos: (user: IEmployee) => void;
    presenceStatus?: string;
    title?: string;
    enableKudos?: boolean;
}

export const ProfileView: React.FunctionComponent<IProfileViewProps> = (props) => {
    const [photoUrl, setPhotoUrl] = React.useState<string>('');
    const [relevantPeople, setRelevantPeople] = React.useState<IEmployee[]>([]);
    const [activeTab, setActiveTab] = React.useState<string>('about');
    const [fullEmployee, setFullEmployee] = React.useState<IEmployee>(props.employee);
    const employeeEmail = props.employee.mail || props.employee.userPrincipalName || '';
    const kudosCount = props.kudosCounts[employeeEmail] || 0;

    React.useEffect(() => {
        setFullEmployee(props.employee);
        const loadData = async () => {
            if (props.employee.id) {
                try {
                    const [photo, , people, details] = await Promise.all([
                        props.graphService.getUserPhoto(props.employee.id),
                        props.graphService.getManagers(props.employee.id),
                        props.graphService.getPeopleIWorkWith(props.employee.id),
                        props.graphService.getUserDetails(props.employee.id)
                    ]);
                    if (photo) setPhotoUrl(photo);
                    setRelevantPeople(people);
                    setFullEmployee(prev => ({ ...prev, ...details }));
                } catch (e) { console.error(e); }
            }
        };
        loadData();
    }, [props.employee.id]);

    const _renderInfoItem = (label: string, value?: string) => (
        <div className={styles.infoItem}>
            <label>{label}</label>
            <span>{value || '---'}</span>
        </div>
    );

    const _renderContactCard = () => (
        <div className={styles.contentCard}>
            <div className={styles.cardTitle}>
                <span><Icon iconName="Contact" /></span>
                <h3>Contact Information</h3>
            </div>
            <div className={styles.infoGrid}>
                {_renderInfoItem('Email', fullEmployee.mail)}
                {_renderInfoItem('Work Phone', fullEmployee.businessPhones?.[0])}
                {_renderInfoItem('Mobile', fullEmployee.mobilePhone)}
                {_renderInfoItem('Office Location', fullEmployee.officeLocation)}
            </div>
        </div>
    );

    const _renderWorkCard = () => (
        <div className={styles.contentCard}>
            <div className={styles.cardTitle}>
                <span><Icon iconName="Org" /></span>
                <h3>Work Information</h3>
            </div>
            <div className={styles.infoGrid}>
                {_renderInfoItem('Department', fullEmployee.department)}
                {_renderInfoItem('Job Title', fullEmployee.jobTitle)}
                {_renderInfoItem('Manager', fullEmployee.manager?.displayName)}
                {_renderInfoItem('Hire Date', fullEmployee.hireDate ? new Date(fullEmployee.hireDate).toLocaleDateString() : undefined)}
            </div>
        </div>
    );

    const _renderBioSection = () => (
        fullEmployee.aboutMe ? (
            <div className={styles.contentCard} style={{ gridColumn: 'span 3' }}>
                <div className={styles.cardTitle}>
                    <span><Icon iconName="ContactInfo" /></span>
                    <h3>About Me</h3>
                </div>
                <div style={{ lineHeight: '1.6', color: 'var(--m-text-dark)', fontSize: '15px' }}>
                    {fullEmployee.aboutMe}
                </div>
            </div>
        ) : null
    );

    const _renderSkillsCard = () => (
        <div className={styles.contentCard}>
            <div className={styles.cardTitle}>
                <span><Icon iconName="Trophy2" /></span>
                <h3>Skills</h3>
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                {fullEmployee.skills && fullEmployee.skills.length > 0 ? (
                    fullEmployee.skills.map((skill, index) => (
                        <span key={`${skill}-${index}`} style={{ background: 'var(--m-bg-hover)', padding: '4px 12px', borderRadius: '16px', fontSize: '12px', fontWeight: 600 }}>{skill}</span>
                    ))
                ) : (
                    <span style={{ color: 'var(--m-text-light)', fontStyle: 'italic' }}>No skills listed</span>
                )}
            </div>
        </div>
    );

    const _renderInterestsCard = () => (
        <div className={styles.contentCard}>
            <div className={styles.cardTitle}>
                <span><Icon iconName="Lightbulb" /></span>
                <h3>Interests</h3>
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                {fullEmployee.interests && fullEmployee.interests.length > 0 ? (
                    fullEmployee.interests.map((interest, index) => (
                        <span key={`${interest}-${index}`} style={{ background: 'var(--m-bg-hover)', padding: '4px 12px', borderRadius: '16px', fontSize: '12px', fontWeight: 600 }}>{interest}</span>
                    ))
                ) : (
                    <span style={{ color: 'var(--m-text-light)', fontStyle: 'italic' }}>No interests listed</span>
                )}
            </div>
        </div>
    );

    const _renderPeopleCard = () => (
        <div className={styles.contentCard}>
            <div className={styles.cardTitle}>
                <span><Icon iconName="People" /></span>
                <h3>People I Work With</h3>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {relevantPeople && relevantPeople.length > 0 ? (
                    relevantPeople.slice(0, 5).map(person => (
                        <button
                            key={person.id}
                            style={{
                                background: 'none', border: 'none', padding: 0, font: 'inherit', cursor: 'pointer', textAlign: 'left',
                                display: 'flex', alignItems: 'center', gap: '12px'
                            }}
                            onClick={() => props.onUserSelect(person)}
                        >
                            <div style={{ width: '32px', height: '32px', borderRadius: '50%', background: 'var(--m-primary-blue)', color: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 600 }}>
                                {person.displayName?.charAt(0)}
                            </div>
                            <div style={{ overflow: 'hidden' }}>
                                <div style={{ fontWeight: 600, fontSize: '14px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{person.displayName}</div>
                                <div style={{ fontSize: '12px', color: 'var(--m-text-medium)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{person.jobTitle}</div>
                            </div>
                        </button>
                    ))
                ) : (
                    <span style={{ color: 'var(--m-text-light)', fontStyle: 'italic' }}>No relevant people found</span>
                )}
            </div>
        </div>
    );

    const _renderOrgChart = () => (
        <div className={styles.contentCard} style={{ gridColumn: 'span 3' }}>
            <div className={styles.cardTitle}>
                <span><Icon iconName="Org" /></span>
                <h3>Organization Chart</h3>
            </div>
            <div style={{ background: '#faf9f8', borderRadius: '8px', padding: '24px', minHeight: '400px' }}>
                <OrgChart currentUser={fullEmployee} graphService={props.graphService} onUserSelect={props.onUserSelect} />
            </div>
        </div>
    );

    const _renderPresenceDot = () => {
        const status = props.presenceStatus || 'Offline';
        const color = status === 'Available' ? 'var(--m-online-green)' : '#ccc';
        return (
            <div style={{
                width: '18px', height: '18px', borderRadius: '50%', background: color, border: '3px solid white',
                position: 'absolute', bottom: '5px', right: '5px', zIndex: 2
            }}></div>
        );
    };

    const _renderWideProfileHeader = () => (
        <div className={styles.contentCard} style={{ gridColumn: 'span 3', padding: '32px', border: '2px solid var(--m-primary-blue) !important', display: 'flex', gap: '0', overflow: 'hidden' }}>
            <div className={styles.cardContentLeft} style={{ flex: '3', paddingRight: '32px' }}>
                <div style={{ flexShrink: 0, position: 'relative' }}>
                    <div style={{ width: '120px', height: '120px', borderRadius: '50%', overflow: 'hidden', border: '4px solid var(--m-bg-light)', boxShadow: 'var(--m-shadow-md)', position: 'relative' }}>
                        {photoUrl ? (
                            <img src={photoUrl} alt={fullEmployee.displayName} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                        ) : (
                            <div style={{ width: '100%', height: '100%', background: 'var(--m-primary-blue)', color: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '56px' }}>
                                {fullEmployee.displayName?.charAt(0)}
                            </div>
                        )}
                    </div>
                    {_renderPresenceDot()}
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <h1 style={{ margin: '0 0 4px 0', fontSize: '28px', color: 'var(--m-text-dark)', fontWeight: 700 }}>{fullEmployee.displayName}</h1>
                    <div style={{ fontSize: '18px', color: 'var(--m-primary-blue)', marginBottom: '8px', fontWeight: 600 }}>{fullEmployee.jobTitle}</div>

                    <div style={{ fontSize: '14px', color: 'var(--m-text-medium)', display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
                        <Icon iconName="Mail" /> <span>{fullEmployee.mail}</span>
                    </div>

                    {fullEmployee.department && (
                        <div style={{ fontSize: '14px', color: 'var(--m-text-medium)', opacity: 0.8 }}>{fullEmployee.department}</div>
                    )}
                </div>
            </div>

            <div className={styles.cardContentRight} style={{ flex: '1', borderLeft: '1px solid #f0f0f0', paddingLeft: '32px' }}>
                <div className={styles.verticalActionList}>
                    <div
                        className={styles.verticalActionItem}
                        onClick={() => globalThis.location.href = `mailto:${fullEmployee.mail}`}
                    >
                        <Icon iconName="Mail" />
                        <span>Send Email</span>
                    </div>
                    <div
                        className={styles.verticalActionItem}
                        onClick={() => globalThis.open(`https://teams.microsoft.com/l/chat/0/0?users=${fullEmployee.mail}`, '_blank')}
                    >
                        <Icon iconName="TeamsLogo" />
                        <span>Teams Chat</span>
                    </div>
                    {props.enableKudos && (
                        <div
                            className={`${styles.verticalActionItem} ${styles.kudosAction}`}
                            onClick={() => props.onGiveKudos(fullEmployee)}
                        >
                            <Icon iconName="FavoriteStarFill" />
                            <span>Kudos {kudosCount > 0 ? `(${kudosCount})` : ''}</span>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );

    const isTabbed = props.visualTheme === 'modern-tabbed-premium';

    return (
        <div className={styles.profileViewContainer}>
            <div className={styles.topBar}>
                <div className={styles.container} style={{ display: 'flex', width: '100%', justifyContent: 'flex-end' }}>
                    <button
                        className={styles.backToDirectoryBtn}
                        onClick={props.onBackToHome}
                    >
                        <Icon iconName="ArrowLeft" /> <span>Back to {props.title || 'Directory'}</span>
                    </button>
                </div>
            </div>

            <div className={styles.container} style={{ marginTop: '0px', width: '100%' }}>
                {isTabbed ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        <div className={styles.contentGrid}>
                            {_renderWideProfileHeader()}
                        </div>

                        <div>
                            <div className={styles.tabTabsList} style={{ display: 'flex', gap: '32px', borderBottom: '1px solid #eee', marginBottom: '32px' }}>
                                <button type="button" className={`${styles.tabTabBtn} ${activeTab === 'about' ? styles.active : ''}`} onClick={() => setActiveTab('about')}>Overview</button>
                                <button type="button" className={`${styles.tabTabBtn} ${activeTab === 'skills' ? styles.active : ''}`} onClick={() => setActiveTab('skills')}>Skills</button>
                                <button type="button" className={`${styles.tabTabBtn} ${activeTab === 'projects' ? styles.active : ''}`} onClick={() => setActiveTab('projects')}>Projects</button>
                                <button type="button" className={`${styles.tabTabBtn} ${activeTab === 'org' ? styles.active : ''}`} onClick={() => setActiveTab('org')}>Organization</button>
                            </div>

                            {activeTab === 'about' && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                                    {_renderBioSection()}
                                    <div className={styles.contentGrid}>
                                        {_renderContactCard()}
                                        {_renderPeopleCard()}
                                    </div>
                                </div>
                            )}

                            {activeTab === 'skills' && (
                                <div className={styles.contentGrid}>
                                    {_renderSkillsCard()}
                                    {_renderInterestsCard()}
                                </div>
                            )}

                            {activeTab === 'projects' && (
                                <div className={styles.contentGrid}>
                                    <div className={styles.contentCard} style={{ gridColumn: 'span 3' }}>
                                        <div className={styles.cardTitle}><h3>Projects</h3></div>
                                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                            {fullEmployee.pastProjects && fullEmployee.pastProjects.length > 0 ? (
                                                fullEmployee.pastProjects.map((p, i) => (
                                                    <span key={`${p}-${i}`} style={{ background: 'var(--m-bg-hover)', padding: '4px 12px', borderRadius: '16px', fontSize: '12px', fontWeight: 600 }}>{p}</span>
                                                ))
                                            ) : <span style={{ fontStyle: 'italic', color: '#888' }}>No projects listed</span>}
                                        </div>
                                    </div>
                                </div>
                            )}

                            {activeTab === 'org' && (
                                <div className={styles.contentGrid}>
                                    {_renderOrgChart()}
                                </div>
                            )}
                        </div>
                    </div>
                ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        <div className={styles.contentGrid}>
                            {_renderWideProfileHeader()}
                        </div>
                        {_renderBioSection()}
                        <div className={styles.contentGrid}>
                            {_renderContactCard()}
                            {_renderWorkCard()}
                            {_renderPeopleCard()}
                        </div>
                        <div className={styles.contentGrid}>
                            {_renderOrgChart()}
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};
