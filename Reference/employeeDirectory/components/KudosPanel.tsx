import * as React from 'react';
import {
    Panel,
    PanelType,
    TextField,
    Dropdown,
    IDropdownOption,
    PrimaryButton,
    DefaultButton,
    Stack,
    Persona,
    PersonaSize,
    Icon,
    Separator,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize
} from '@fluentui/react';
import { IKudos, KudosService } from '../services/KudosService';
import { IEmployee } from '../models/IEmployee';
import styles from './EmployeeDirectory.module.scss';

export interface IKudosPanelProps {
    isOpen: boolean;
    onDismiss: () => void;
    mode: 'give' | 'view';
    targetUser: IEmployee | null;
    currentUser: IEmployee | undefined;
    kudosService: KudosService | undefined;
    onKudosSubmitted?: () => void;
    themeClass?: string;
}

export const KudosPanel: React.FC<IKudosPanelProps> = (props) => {
    const [message, setMessage] = React.useState('');
    const [kudosType, setKudosType] = React.useState<string>('Excellence');
    const [isSubmitting, setIsSubmitting] = React.useState(false);
    const [status, setStatus] = React.useState<{ text: string, type: MessageBarType } | null>(null);
    const [kudosList, setKudosList] = React.useState<IKudos[]>([]);
    const [isLoading, setIsLoading] = React.useState(false);
    const [isFormVisible, setIsFormVisible] = React.useState(props.mode === 'give');
    const [photoUrl, setPhotoUrl] = React.useState<string>('');

    const kudosTypeOptions: IDropdownOption[] = [
        { key: 'Excellence', text: 'Excellence' },
        { key: 'Innovation', text: 'Innovation' },
        { key: 'Teamwork', text: 'Teamwork' },
        { key: 'Leadership', text: 'Leadership' },
        { key: 'Customer Focus', text: 'Customer Focus' },
        { key: 'Going Above & Beyond', text: 'Going Above & Beyond' }
    ];

    React.useEffect(() => {
        if (props.isOpen && props.targetUser) {
            _fetchKudos();
            _fetchPhoto();
            setIsFormVisible(props.mode === 'give');
        }
        if (!props.isOpen) {
            setMessage('');
            setStatus(null);
        }
    }, [props.isOpen, props.targetUser, props.mode]);

    const _fetchPhoto = () => {
        if (!props.targetUser) return;
        setPhotoUrl(`/_layouts/15/userphoto.aspx?size=L&accountname=${props.targetUser.mail || props.targetUser.userPrincipalName}`);
    };

    const _fetchKudos = async () => {
        if (!props.targetUser || !props.kudosService) return;
        setIsLoading(true);
        const email = props.targetUser.mail || props.targetUser.userPrincipalName || '';
        try {
            const list = await props.kudosService.getKudosForUser(email);
            setKudosList(list);
        } catch (err) {
            console.error("Error fetching kudos", err);
        } finally {
            setIsLoading(false);
        }
    };

    const _handleSubmit = async () => {
        if (!props.targetUser || !props.kudosService || !props.currentUser) return;

        setIsSubmitting(true);
        const email = props.targetUser.mail || props.targetUser.userPrincipalName || '';
        const senderEmail = props.currentUser.mail || props.currentUser.userPrincipalName || '';
        const success = await props.kudosService.submitKudos({
            title: `Kudos for ${props.targetUser.displayName}`,
            targetUserEmail: email,
            message: message,
            kudosType: kudosType,
            senderName: props.currentUser.displayName || '',
            senderEmail: senderEmail
        });

        if (success) {
            setStatus({ text: 'Kudos submitted successfully!', type: MessageBarType.success });
            setMessage('');
            setIsFormVisible(false);
            _fetchKudos();
            if (props.onKudosSubmitted) props.onKudosSubmitted();
            setTimeout(() => setStatus(null), 3000);
        } else {
            setStatus({ text: 'Error submitting kudos. Please try again.', type: MessageBarType.error });
        }
        setIsSubmitting(false);
    };

    if (!props.targetUser) return null;

    const _renderKudosList = () => {
        if (isLoading) {
            return (
                <div style={{ padding: '40px 0', textAlign: 'center' }}>
                    <Spinner size={SpinnerSize.large} label="Loading recognition history..." />
                </div>
            );
        }

        if (kudosList.length > 0) {
            return (
                <div className={styles.kudosList}>
                    {kudosList.map((k) => (
                        <div key={k.id} className={styles.kudosItem}>
                            <div className={styles.kudosHeader}>
                                <div className={styles.type}>
                                    <Icon iconName="FavoriteStarFill" /> {k.kudosType}
                                </div>
                                <div className={styles.date}>
                                    {k.kudosDate ? new Date(k.kudosDate).toLocaleString() : 'Recent'}
                                </div>
                            </div>
                            <div className={styles.kudosMessage}>{k.message}</div>
                            <div className={styles.kudosFrom}>
                                <Persona
                                    text={k.senderName}
                                    size={PersonaSize.size24}
                                />
                            </div>
                        </div>
                    ))}
                </div>
            );
        }

        return (
            <div style={{ textAlign: 'center', padding: '40px 20px', opacity: 0.6 }}>
                <Icon iconName="Ribbon" style={{ fontSize: 48, marginBottom: 15, color: '#eee' }} />
                <div>No recognition history yet.</div>
                <div style={{ fontSize: '12px', marginTop: '5px' }}>Be the first to share some appreciation!</div>
            </div>
        );
    };

    return (
        <React.Fragment>
            {/* Custom Side Panel (Kudos) */}
            <button
                className={`${styles.overlay} ${props.isOpen ? styles.open : ''}`}
                onClick={props.onDismiss}
                style={{
                    border: 'none', padding: 0, margin: 0, cursor: 'default',
                    position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh',
                    background: 'rgba(0,0,0,0.3)', zIndex: 999,
                    display: props.isOpen ? 'block' : 'none'
                }}
                aria-label="Close kudos panel"
            ></button>
            <div className={`${styles.sidePanel} ${props.isOpen ? styles.open : ''}`}>
                <div className={styles.sidePanelHeader}>
                    <h2>Kudos</h2>
                    <button className={styles.closePanel} onClick={props.onDismiss}>✕</button>
                </div>
                <div className={styles.sidePanelContent} style={{ padding: '0 24px 24px' }}>
                    <div className={styles.kudosPanelInner}>
                        {/* 1. Identity Header */}
                        <div className={styles.panelEmployeeInfo} style={{ display: 'flex', alignItems: 'center', gap: '20px', padding: '10px 0' }}>
                            <div className={styles.panelAvatar} style={{ width: '64px', height: '64px', border: '2px solid #eee' }}>
                                <Persona
                                    imageUrl={photoUrl}
                                    size={PersonaSize.size56}
                                    hidePersonaDetails
                                />
                            </div>
                            <div className={styles.panelInfoText}>
                                <div className={styles.panelEmployeeName} style={{ fontSize: '24px', fontWeight: 600 }}>{props.targetUser.displayName}</div>
                                <div className={styles.panelEmployeeTitle}>{props.targetUser.jobTitle}</div>
                            </div>
                        </div>

                        {/* 2. Status Messages */}
                        {status && (
                            <MessageBar
                                messageBarType={status.type}
                                onDismiss={() => setStatus(null)}
                                style={{ marginBottom: 20 }}
                            >
                                {status.text}
                            </MessageBar>
                        )}

                        {/* 3. Give Kudos Section (Button or Form) */}
                        {isFormVisible ? (
                            <div className={styles.kudosForm}>
                                <h3 className={styles.kudosSectionTitle} style={{ marginBottom: '15px' }}>
                                    <Icon iconName="FavoriteStar" /> Give Recognition
                                </h3>
                                <Stack tokens={{ childrenGap: 20 }}>
                                    <Dropdown
                                        label="Kudos Type"
                                        selectedKey={kudosType}
                                        options={kudosTypeOptions}
                                        onChange={(_, opt) => setKudosType(opt?.key as string)}
                                        className={styles.formSelect}
                                    />
                                    <TextField
                                        label="Message"
                                        multiline
                                        rows={4}
                                        placeholder="Why are you giving kudos?"
                                        value={message}
                                        onChange={(_, val) => setMessage(val || '')}
                                        required
                                        className={styles.formTextarea}
                                    />
                                    <div className={styles.formActions} style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
                                        <PrimaryButton
                                            onClick={_handleSubmit}
                                            disabled={isSubmitting || !message.trim()}
                                            className={styles.btnSubmit}
                                        >
                                            {isSubmitting ? <Spinner size={SpinnerSize.xSmall} style={{ marginRight: 8 }} /> : null}
                                            Submit
                                        </PrimaryButton>
                                        <DefaultButton
                                            text="Cancel"
                                            onClick={() => setIsFormVisible(false)}
                                            className={styles.btnCancel}
                                        />
                                    </div>
                                </Stack>
                                <Separator />
                            </div>
                        ) : (
                            <div
                                className={styles.giveKudosBtn}
                                onClick={() => setIsFormVisible(true)}
                                role="button"
                                style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                            >
                                <span style={{ fontSize: '20px', marginRight: '8px' }}>⭐</span> Give Kudos
                            </div>
                        )}

                        {/* 4. Kudos History Section */}
                        <div className={styles.kudosListSection}>
                            <h3 className={styles.kudosListTitle}>
                                Kudos Received ({kudosList.length})
                            </h3>
                            {_renderKudosList()}
                        </div>
                    </div>
                </div>
            </div>
        </React.Fragment>
    );


};
