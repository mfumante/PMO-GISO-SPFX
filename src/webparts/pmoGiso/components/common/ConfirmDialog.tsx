import * as React from 'react';
import {
  Button,
  CircularProgress,
  Dialog,
  DialogActions,
  DialogContent,
  DialogContentText,
  DialogTitle,
} from '@mui/material';

export interface IConfirmDialogProps {
  open: boolean;
  title: string;
  message: React.ReactNode;
  confirmLabel?: string;
  confirmColor?: 'error' | 'warning' | 'primary';
  loading?: boolean;
  onCancel: () => void;
  onConfirm: () => void;
}

// Dialog di conferma generico, riusato per tutte le eliminazioni (progetti,
// deliverable, issue, costi, allocazioni, risorse).
export function ConfirmDialog({
  open,
  title,
  message,
  confirmLabel,
  confirmColor,
  loading,
  onCancel,
  onConfirm,
}: IConfirmDialogProps): React.ReactElement {
  return (
    <Dialog open={open} onClose={loading ? undefined : onCancel} fullWidth maxWidth="xs">
      <DialogTitle>{title}</DialogTitle>
      <DialogContent>
        <DialogContentText component="div">{message}</DialogContentText>
      </DialogContent>
      <DialogActions>
        <Button onClick={onCancel} disabled={loading}>
          Annulla
        </Button>
        <Button
          variant="contained"
          color={confirmColor ?? 'error'}
          onClick={onConfirm}
          disabled={loading}
          startIcon={loading ? <CircularProgress size={16} color="inherit" /> : undefined}
        >
          {confirmLabel ?? 'Elimina'}
        </Button>
      </DialogActions>
    </Dialog>
  );
}
