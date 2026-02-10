import * as React from 'react';
import styles from './GiveKudosForm.module.scss';

export interface IGiveKudosFormProps {
    recipientName: string;
    badgeChoices: string[]; // Dynamic choices from SharePoint
    onSubmit: (message: string, badgeType: string) => Promise<void>;
    onCancel: () => void;
}

export const GiveKudosForm: React.FC<IGiveKudosFormProps> = ({
    recipientName,
    badgeChoices,
    onSubmit,
    onCancel
}) => {
    const [message, setMessage] = React.useState('');
    const [selectedBadge, setSelectedBadge] = React.useState(badgeChoices[0] || '');
    const [submitting, setSubmitting] = React.useState(false);
    const [error, setError] = React.useState('');
    const [success, setSuccess] = React.useState(false);

    const handleSubmit = async (): Promise<void> => {
        if (!message.trim()) {
            setError('Please enter a message');
            return;
        }

        setSubmitting(true);
        setError('');
        setSuccess(false);

        try {
            await onSubmit(message, selectedBadge);
            setSuccess(true);
            // Form will be closed by parent component after a brief delay
            setTimeout(() => {
                onCancel(); // Close the form
            }, 1500);
        } catch (err) {
            console.error('[GiveKudosForm] Error giving kudos:', err);
            setError('Failed to send kudos. Please try again.');
            setSubmitting(false);
        }
    };

    return (
        <div className={styles.giveKudosForm}>
            <div className={styles.header}>
                <h2 className={styles.title}>Give Kudos</h2>
                <button className={styles.closeButton} onClick={onCancel} aria-label="Close">
                    ×
                </button>
            </div>

            {error && <div className={styles.error}>{error}</div>}
            {success && <div className={styles.success}>✅ Kudos sent successfully!</div>}

            <div className={styles.formField}>
                <label htmlFor="kudos-recipient" className={styles.label}>To:</label>
                <div id="kudos-recipient" className={styles.recipientDisplay}>{recipientName}</div>
            </div>

            <div className={styles.formField}>
                <label htmlFor="kudos-message" className={styles.label}>Message:</label>
                <textarea
                    id="kudos-message"
                    className={styles.textarea}
                    value={message}
                    onChange={(e) => setMessage(e.target.value)}
                    placeholder="Share why they're awesome..."
                    disabled={submitting}
                />
            </div>

            <div className={styles.formField}>
                <label htmlFor="kudos-badge" className={styles.label}>Choose a badge:</label>
                <select
                    id="kudos-badge"
                    className={styles.badgeSelect}
                    value={selectedBadge}
                    onChange={(e) => setSelectedBadge(e.target.value)}
                    disabled={submitting}
                >
                    {badgeChoices.map((choice: string) => (
                        <option key={choice} value={choice}>
                            {choice}
                        </option>
                    ))}
                </select>
            </div>

            <div className={styles.actions}>
                <button
                    className={styles.cancelButton}
                    onClick={onCancel}
                    disabled={submitting}
                >
                    Cancel
                </button>
                <button
                    className={styles.submitButton}
                    onClick={handleSubmit}
                    disabled={submitting || !message.trim()}
                >
                    {submitting ? 'Sending...' : 'Send Kudos'}
                </button>
            </div>
        </div>
    );
};
