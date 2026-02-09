import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';

export interface IAlphabetNavProps {
    onLetterClick: (letter: string) => void;
    selectedLetter: string;
    enableStarred?: boolean;
}

export const AlphabetNav: React.FunctionComponent<IAlphabetNavProps> = (props) => {
    const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');

    return (
        <div className={styles.letters}>
            {props.enableStarred && (
                <button
                    className={`${styles.letter} ${styles.star} ${props.selectedLetter === '*' ? styles.active : ''}`}
                    onClick={() => props.onLetterClick('*')}
                    title="Featured"
                    type="button"
                >
                    ⭐
                </button>
            )}
            <button
                className={`${styles.letter} ${props.selectedLetter === '' ? styles.active : ''}`}
                onClick={() => props.onLetterClick('')}
                type="button"
            >
                All
            </button>
            {letters.map(letter => (
                <button
                    key={letter}
                    className={`${styles.letter} ${props.selectedLetter === letter ? styles.active : ''}`}
                    onClick={() => props.onLetterClick(letter)}
                    type="button"
                >
                    {letter}
                </button>
            ))}
        </div>
    );
};
