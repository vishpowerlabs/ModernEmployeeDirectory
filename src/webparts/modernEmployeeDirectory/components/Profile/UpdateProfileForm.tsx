import * as React from 'react';
import styles from './UpdateProfileForm.module.scss';

export interface IUpdateProfileFormProps {
    user: {
        id: string;
        name: string;
        email: string;
        jobTitle?: string;
        aboutMe?: string;
        mobilePhone?: string;
        officeLocation?: string;
        skills?: string[];
        interests?: string[];
        pastProjects?: string[];
    };
    updatableFields: string[];
    onSubmit: (updatedFields: any) => Promise<any>;
    onCancel: () => void;
    loading?: boolean;
}

export type TagField = 'skills' | 'interests' | 'pastProjects';

export const UpdateProfileForm: React.FC<IUpdateProfileFormProps> = ({
    user,
    updatableFields,
    onSubmit,
    onCancel,
    loading = false
}) => {
    const [formData, setFormData] = React.useState<any>({
        jobTitle: user.jobTitle || '',
        aboutMe: user.aboutMe || '',
        mobilePhone: user.mobilePhone || '',
        officeLocation: user.officeLocation || '',
        skills: user.skills || [],
        interests: user.interests || [],
        pastProjects: user.pastProjects || []
    });

    const [skillInput, setSkillInput] = React.useState('');
    const [interestInput, setInterestInput] = React.useState('');
    const [pastProjectInput, setPastProjectInput] = React.useState('');

    const [updatedStatus, setUpdatedStatus] = React.useState<Record<string, boolean>>({});
    const [submitting, setSubmitting] = React.useState(false);
    const [error, setError] = React.useState('');
    const [success, setSuccess] = React.useState(false);

    const handleInputChange = (field: string, value: any): void => {
        setFormData((prev: any) => ({ ...prev, [field]: value }));
    };

    const handleAddTag = (field: TagField, value: string): void => {
        const trimmedValue = value.trim();
        if (trimmedValue && !formData[field].includes(trimmedValue)) {
            const newTags = [...formData[field], trimmedValue];
            handleInputChange(field, newTags);
        }
        if (field === 'skills') setSkillInput('');
        else if (field === 'interests') setInterestInput('');
        else if (field === 'pastProjects') setPastProjectInput('');
    };

    const handleRemoveTag = (field: TagField, tagToRemove: string): void => {
        const newTags = formData[field].filter((tag: string) => tag !== tagToRemove);
        handleInputChange(field, newTags);
    };

    const handleTagKeyPress = (e: React.KeyboardEvent, field: TagField, value: string): void => {
        if (e.key === 'Enter') {
            e.preventDefault();
            handleAddTag(field, value);
        }
    };

    // Sync form data when user prop changes (e.g., after async fetch in parent)
    React.useEffect(() => {
        if (!loading) {
            setFormData({
                jobTitle: user.jobTitle || '',
                aboutMe: user.aboutMe || '',
                mobilePhone: user.mobilePhone || '',
                officeLocation: user.officeLocation || '',
                skills: user.skills || [],
                interests: user.interests || [],
                pastProjects: user.pastProjects || []
            });
        }
    }, [user, loading]);

    const handleSubmit = async (): Promise<void> => {
        setSubmitting(true);
        setError('');
        setSuccess(false);

        try {
            // Prepare data for Graph API
            const updatePayload: any = {};
            const newUpdatedStatus: Record<string, boolean> = { ...updatedStatus };

            updatableFields.forEach(field => {
                updatePayload[field] = formData[field];

                // Track what was changed for "already updated" status
                const originalValue = (user as any)[field];
                const currentValue = formData[field];

                if (Array.isArray(currentValue) && Array.isArray(originalValue)) {
                    // Simple array comparison
                    const isDifferent = currentValue.length !== originalValue.length ||
                        !currentValue.every((v, i) => v === originalValue[i]);
                    if (isDifferent) {
                        newUpdatedStatus[field] = true;
                    }
                } else if (currentValue !== originalValue) {
                    newUpdatedStatus[field] = true;
                }
            });

            const result = await onSubmit(updatePayload);
            if (result === true || (typeof result === 'object' && result.success)) {
                setSuccess(true);
                setUpdatedStatus(newUpdatedStatus);
                setTimeout(() => {
                    setSuccess(false);
                }, 3000);
            } else {
                setError(typeof result === 'object' && result.details ? `Update failed: ${result.details}` : 'Failed to update profile. Please try again.');
            }
        } catch (err: any) {
            setError(`An error occurred: ${err.message || 'Unknown error'}`);
        } finally {
            setSubmitting(false);
        }
    };

    const renderField = (field: string): React.ReactElement | null => {
        const isRecentlyUpdated = updatedStatus[field];

        switch (field) {
            case 'jobTitle':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Job Title {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <input
                            type="text"
                            className={styles.input}
                            value={formData.jobTitle}
                            onChange={(e) => handleInputChange('jobTitle', e.target.value)}
                            disabled={submitting}
                        />
                    </div>
                );
            case 'aboutMe':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Bio {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <textarea
                            className={styles.textarea}
                            value={formData.aboutMe}
                            onChange={(e) => handleInputChange('aboutMe', e.target.value)}
                            disabled={submitting}
                        />
                    </div>
                );
            case 'mobilePhone':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Mobile Phone {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <input
                            type="text"
                            className={styles.input}
                            value={formData.mobilePhone}
                            onChange={(e) => handleInputChange('mobilePhone', e.target.value)}
                            disabled={submitting}
                        />
                    </div>
                );
            case 'officeLocation':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Office Location {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <input
                            type="text"
                            className={styles.input}
                            value={formData.officeLocation}
                            onChange={(e) => handleInputChange('officeLocation', e.target.value)}
                            disabled={submitting}
                        />
                    </div>
                );
            case 'skills':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Skills {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <div className={styles.tagPicker}>
                            {formData.skills.map((tag: string) => (
                                <div key={tag} className={styles.tagItem}>
                                    <span className={styles.tagText}>{tag}</span>
                                    <button
                                        type="button"
                                        className={styles.removeTag}
                                        onClick={() => handleRemoveTag('skills', tag)}
                                        onKeyDown={(e) => {
                                            if (e.key === 'Enter' || e.key === ' ') {
                                                e.preventDefault();
                                                handleRemoveTag('skills', tag);
                                            }
                                        }}
                                        aria-label={`Remove ${tag}`}
                                        disabled={submitting}
                                    >×</button>
                                </div>
                            ))}
                            <input
                                type="text"
                                className={styles.tagInput}
                                placeholder="Add skill..."
                                value={skillInput}
                                onChange={(e) => setSkillInput(e.target.value)}
                                onKeyDown={(e) => handleTagKeyPress(e, 'skills', skillInput)}
                                onBlur={() => handleAddTag('skills', skillInput)}
                                disabled={submitting}
                            />
                        </div>
                    </div>
                );
            case 'interests':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Interests {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <div className={styles.tagPicker}>
                            {formData.interests.map((tag: string) => (
                                <div key={tag} className={styles.tagItem}>
                                    <span className={styles.tagText}>{tag}</span>
                                    <button
                                        type="button"
                                        className={styles.removeTag}
                                        onClick={() => handleRemoveTag('interests', tag)}
                                        onKeyDown={(e) => {
                                            if (e.key === 'Enter' || e.key === ' ') {
                                                e.preventDefault();
                                                handleRemoveTag('interests', tag);
                                            }
                                        }}
                                        aria-label={`Remove ${tag}`}
                                        disabled={submitting}
                                    >×</button>
                                </div>
                            ))}
                            <input
                                type="text"
                                className={styles.tagInput}
                                placeholder="Add interest..."
                                value={interestInput}
                                onKeyDown={(e) => handleTagKeyPress(e, 'interests', interestInput)}
                                onBlur={() => handleAddTag('interests', interestInput)}
                                disabled={submitting}
                            />
                        </div>
                    </div>
                );

            case 'pastProjects':
                return (
                    <div className={styles.formField} key={field}>
                        <label className={styles.label}>
                            Past Projects {isRecentlyUpdated && <span className={styles.recentUpdate}>(Recently Updated)</span>}
                        </label>
                        <div className={styles.tagPicker}>
                            {formData.pastProjects.map((tag: string) => (
                                <div key={tag} className={styles.tagItem}>
                                    <span className={styles.tagText}>{tag}</span>
                                    <button
                                        type="button"
                                        className={styles.removeTag}
                                        onClick={() => handleRemoveTag('pastProjects', tag)}
                                        onKeyDown={(e) => {
                                            if (e.key === 'Enter' || e.key === ' ') {
                                                e.preventDefault();
                                                handleRemoveTag('pastProjects', tag);
                                            }
                                        }}
                                        aria-label={`Remove ${tag}`}
                                        disabled={submitting}
                                    >×</button>
                                </div>
                            ))}
                            <input
                                type="text"
                                className={styles.tagInput}
                                placeholder="Add past project..."
                                value={pastProjectInput}
                                onChange={(e) => setPastProjectInput(e.target.value)}
                                onKeyDown={(e) => handleTagKeyPress(e, 'pastProjects', pastProjectInput)}
                                onBlur={() => handleAddTag('pastProjects', pastProjectInput)}
                                disabled={submitting}
                            />
                        </div>
                    </div>
                );
            default:
                return null;
        }
    };

    return (
        <div className={styles.updateProfileForm}>
            {loading && (
                <div className={styles.loadingOverlay}>
                    <div className={styles.spinner}></div>
                    <p>Fetching latest profile data...</p>
                </div>
            )}
            <div className={styles.header}>
                <h2 className={styles.title}>Update Profile</h2>
                <button className={styles.closeButton} onClick={onCancel}>×</button>
            </div>

            {error && <div className={styles.error}>{error}</div>}
            {success && <div className={styles.success}>✅ Profile updated successfully!</div>}

            <div style={{ flex: 1, overflowY: 'auto' }}>
                {updatableFields.length === 0 ? (
                    <div className={styles.error}>No fields configured for update.</div>
                ) : (
                    updatableFields.map(field => renderField(field))
                )}
            </div>

            <div className={styles.actions}>
                <button className={styles.cancelButton} onClick={onCancel} disabled={submitting}>
                    Cancel
                </button>
                <button
                    className={styles.submitButton}
                    onClick={handleSubmit}
                    disabled={submitting || updatableFields.length === 0}
                >
                    {submitting ? 'Updating...' : 'Save Changes'}
                </button>
            </div>
        </div>
    );
};
