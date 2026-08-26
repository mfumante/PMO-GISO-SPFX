import { useState } from 'react';

export interface IFeedbackState {
  open: boolean;
  message: string;
  severity: 'success' | 'error';
}

export interface IFeedbackControls {
  feedback: IFeedbackState;
  showSuccess: (message: string) => void;
  showError: (message: string) => void;
  closeFeedback: () => void;
}

const INITIAL_STATE: IFeedbackState = { open: false, message: '', severity: 'success' };

// Stato dello Snackbar di conferma/errore, condiviso da Portfolio, Scheda Progetto
// e Risorse dopo ogni operazione di creazione/modifica/eliminazione.
export function useFeedback(): IFeedbackControls {
  const [feedback, setFeedback] = useState<IFeedbackState>(INITIAL_STATE);

  function showSuccess(message: string): void {
    setFeedback({ open: true, message, severity: 'success' });
  }

  function showError(message: string): void {
    setFeedback({ open: true, message, severity: 'error' });
  }

  function closeFeedback(): void {
    setFeedback((prev) => ({ ...prev, open: false }));
  }

  return { feedback, showSuccess, showError, closeFeedback };
}
