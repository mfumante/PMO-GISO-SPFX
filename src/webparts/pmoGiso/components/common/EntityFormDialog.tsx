import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Alert,
  Button,
  CircularProgress,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  FormControl,
  FormControlLabel,
  FormHelperText,
  InputLabel,
  MenuItem,
  Select,
  Stack,
  Switch,
  TextField,
  Typography,
} from '@mui/material';
import type { SelectChangeEvent } from '@mui/material';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IFieldHelp } from '../help/fieldHelp';
import { HelpLabel } from '../help/HelpLabel';
import { IPersonValue, PeoplePicker } from './PeoplePicker';

export type FormFieldType = 'text' | 'note' | 'date' | 'number' | 'currency' | 'boolean' | 'select' | 'person' | 'hidden';

export interface IFormFieldOption {
  value: string;
  label: string;
}

export interface IFormFieldSchema {
  name: string;
  label: string;
  type: FormFieldType;
  required?: boolean;
  disabled?: boolean;
  options?: IFormFieldOption[];
  min?: number;
  max?: number;
  // Guida contestuale: short va sotto il campo come helperText, long+example
  // nel tooltip dell'icona informativa accanto alla label. Vedi fieldHelp.ts.
  help?: IFieldHelp;
  // Solo per campi 'person': nome del campo testo da valorizzare con il display
  // name quando l'utente seleziona una persona (es. 'Sponsor' per 'SponsorUser'),
  // cosi' le due colonne restano allineate senza sostituire il campo testo.
  syncTextField?: string;
  // Solo per campi 'person': nome del campo (tipicamente 'hidden') da valorizzare
  // con l'email della persona selezionata, per storicizzarla nella lista anche se
  // l'identita' cambia o l'utente lascia il tenant. Come syncTextField, si aggiorna
  // solo alla selezione: svuotare il campo persona lascia l'email gia' registrata.
  syncEmailField?: string;
}

export type FormFieldValue = string | number | boolean | IPersonValue | undefined;
export type FormValues = Record<string, FormFieldValue>;

export interface IDateRangeRule {
  startField: string;
  endField: string;
  message?: string;
}

export interface IEntityFormDialogProps {
  context: WebPartContext;
  open: boolean;
  title: string;
  fields: IFormFieldSchema[];
  initialValues: FormValues;
  dateRange?: IDateRangeRule;
  onCancel: () => void;
  onSave: (values: FormValues) => Promise<void>;
}

function isPersonValue(value: FormFieldValue): value is IPersonValue {
  return !!value && typeof value === 'object';
}

// Converte un valore data (stringa ISO letta da SharePoint, o gia' in formato
// gg del browser) nel formato 'YYYY-MM-DD' richiesto da <input type="date">.
function toDateInputValue(value: FormFieldValue): string {
  if (!value || typeof value !== 'string') {
    return '';
  }
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return '';
  }
  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  return `${year}-${month}-${day}`;
}

// Stato di editing: tutti i campi sono tenuti come stringa (tranne boolean e
// person), coerente con i controlli HTML che li rendono (TextField/Select).
function buildEditValues(fields: IFormFieldSchema[], initialValues: FormValues): FormValues {
  const values: FormValues = {};
  fields.forEach((field) => {
    const raw = initialValues[field.name];
    if (field.type === 'boolean') {
      values[field.name] = raw === true;
    } else if (field.type === 'person') {
      values[field.name] = isPersonValue(raw) ? raw : undefined;
    } else if (field.type === 'date') {
      values[field.name] = toDateInputValue(raw);
    } else {
      values[field.name] = raw == null ? '' : String(raw);
    }
  });
  return values;
}

// Label composta da testo + icona informativa (vedi HelpLabel): usata sia per
// TextField/Switch (il loro prop 'label' accetta un ReactNode) sia per
// InputLabel dei Select.
function renderFieldLabel(label: string, help: IFieldHelp | undefined): React.ReactNode {
  return <HelpLabel label={label} help={help} />;
}

// Componente Dialog generico per la creazione/modifica di un record: i campi
// vengono generati dallo schema passato dal chiamante (vedi entityFormSchemas.ts),
// cosi' la stessa UI serve per tutte le entita' (Project, Deliverable, Issue,
// Resource, Allocation, Cost) senza duplicare markup o logica di validazione.
export function EntityFormDialog({
  context,
  open,
  title,
  fields,
  initialValues,
  dateRange,
  onCancel,
  onSave,
}: IEntityFormDialogProps): React.ReactElement {
  const [values, setValues] = useState<FormValues>(() => buildEditValues(fields, initialValues));
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [formError, setFormError] = useState<string | undefined>(undefined);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (open) {
      setValues(buildEditValues(fields, initialValues));
      setFieldErrors({});
      setFormError(undefined);
      setSaving(false);
    }
    // Si vuole risincronizzare solo all'apertura del dialog, non ad ogni
    // variazione di riferimento di 'fields'/'initialValues' durante l'editing.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  function setFieldValue(name: string, value: FormFieldValue): void {
    setValues((prev) => ({ ...prev, [name]: value }));
  }

  // Selezionare una persona valorizza anche il campo testo collegato (mantenendo
  // i due dati allineati); svuotare la selezione lascia il testo com'e', cosi' chi
  // vuole registrare un soggetto esterno al tenant puo' comunque digitarlo a mano.
  function handlePersonChange(field: IFormFieldSchema, person: IPersonValue | undefined): void {
    setFieldValue(field.name, person);
    if (person && field.syncTextField) {
      setFieldValue(field.syncTextField, person.displayName);
    }
    if (person && field.syncEmailField) {
      setFieldValue(field.syncEmailField, person.email);
    }
  }

  function validateAndBuildPayload(): { payload?: FormValues; errors: Record<string, string> } {
    const errors: Record<string, string> = {};
    const payload: FormValues = {};

    fields.forEach((field) => {
      const raw = values[field.name];

      if (field.type === 'boolean') {
        payload[field.name] = raw === true;
        return;
      }

      if (field.type === 'person') {
        payload[field.name] = isPersonValue(raw) ? raw : undefined;
        return;
      }

      if (field.type === 'number' || field.type === 'currency') {
        const text = typeof raw === 'string' ? raw.trim() : '';
        if (!text) {
          if (field.required) {
            errors[field.name] = 'Campo obbligatorio.';
          } else {
            payload[field.name] = undefined;
          }
          return;
        }
        const parsed = Number(text);
        if (isNaN(parsed)) {
          errors[field.name] = 'Deve essere un numero valido.';
          return;
        }
        if (field.min != null && parsed < field.min) {
          errors[field.name] = `Il valore minimo e' ${field.min}.`;
          return;
        }
        if (field.max != null && parsed > field.max) {
          errors[field.name] = `Il valore massimo e' ${field.max}.`;
          return;
        }
        payload[field.name] = parsed;
        return;
      }

      if (field.type === 'date') {
        const text = typeof raw === 'string' ? raw.trim() : '';
        if (!text) {
          if (field.required) {
            errors[field.name] = 'Campo obbligatorio.';
          } else {
            payload[field.name] = undefined;
          }
          return;
        }
        const date = new Date(`${text}T00:00:00`);
        if (isNaN(date.getTime())) {
          errors[field.name] = 'Data non valida.';
          return;
        }
        payload[field.name] = date.toISOString();
        return;
      }

      // text | note | select | hidden
      const text = typeof raw === 'string' ? raw.trim() : '';
      if (!text) {
        if (field.required) {
          errors[field.name] = 'Campo obbligatorio.';
        } else {
          payload[field.name] = undefined;
        }
        return;
      }
      payload[field.name] = text;
    });

    if (Object.keys(errors).length === 0 && dateRange) {
      const start = payload[dateRange.startField];
      const end = payload[dateRange.endField];
      if (typeof start === 'string' && typeof end === 'string' && new Date(end).getTime() < new Date(start).getTime()) {
        errors[dateRange.endField] = dateRange.message ?? 'La data di fine non puo\' essere precedente alla data di inizio.';
      }
    }

    return { payload: Object.keys(errors).length === 0 ? payload : undefined, errors };
  }

  async function handleSave(): Promise<void> {
    const { payload, errors } = validateAndBuildPayload();
    setFieldErrors(errors);
    if (!payload) {
      return;
    }

    setSaving(true);
    setFormError(undefined);
    try {
      await onSave(payload);
    } catch (error) {
      // In caso di errore il dialog resta aperto con i dati inseriti dall'utente,
      // che non deve reinserirli da capo.
      setFormError(error instanceof Error ? error.message : 'Errore durante il salvataggio.');
    } finally {
      setSaving(false);
    }
  }

  function handleClose(): void {
    if (!saving) {
      onCancel();
    }
  }

  return (
    <Dialog open={open} onClose={handleClose} fullWidth maxWidth="sm">
      <DialogTitle>{title}</DialogTitle>
      <DialogContent>
        <Stack spacing={2} sx={{ mt: 1 }}>
          {formError && <Alert severity="error">{formError}</Alert>}
          {fields.map((field) => {
            const value = values[field.name];
            const error = fieldErrors[field.name];
            const fieldLabel = renderFieldLabel(field.label, field.help);

            // Campo 'hidden': valorizzato solo a programma (vedi syncEmailField),
            // mai mostrato all'utente.
            if (field.type === 'hidden') {
              return null;
            }

            if (field.type === 'person') {
              return (
                <PeoplePicker
                  key={field.name}
                  context={context}
                  label={field.label}
                  value={isPersonValue(value) ? value : undefined}
                  disabled={field.disabled || saving}
                  error={!!error}
                  helperText={error ?? field.help?.short}
                  onChange={(person) => handlePersonChange(field, person)}
                />
              );
            }

            if (field.type === 'boolean') {
              return (
                <Stack key={field.name} spacing={0}>
                  <FormControlLabel
                    control={
                      <Switch
                        checked={value === true}
                        onChange={(event) => setFieldValue(field.name, event.target.checked)}
                        disabled={field.disabled || saving}
                      />
                    }
                    label={fieldLabel}
                  />
                  {field.help && (
                    <Typography variant="caption" color="text.secondary">
                      {field.help.short}
                    </Typography>
                  )}
                </Stack>
              );
            }

            if (field.type === 'select') {
              return (
                <FormControl key={field.name} size="small" error={!!error} disabled={field.disabled || saving}>
                  <InputLabel id={`entity-form-field-${field.name}`}>{fieldLabel}</InputLabel>
                  <Select
                    labelId={`entity-form-field-${field.name}`}
                    label={field.label}
                    value={typeof value === 'string' ? value : ''}
                    onChange={(event: SelectChangeEvent) => setFieldValue(field.name, event.target.value)}
                  >
                    {(field.options ?? []).map((option) => (
                      <MenuItem key={option.value} value={option.value}>
                        {option.label}
                      </MenuItem>
                    ))}
                  </Select>
                  {(error || field.help?.short) && <FormHelperText>{error ?? field.help?.short}</FormHelperText>}
                </FormControl>
              );
            }

            return (
              <TextField
                key={field.name}
                label={fieldLabel}
                value={typeof value === 'string' ? value : ''}
                onChange={(event) => setFieldValue(field.name, event.target.value)}
                type={field.type === 'date' ? 'date' : field.type === 'number' || field.type === 'currency' ? 'number' : 'text'}
                multiline={field.type === 'note'}
                minRows={field.type === 'note' ? 3 : undefined}
                error={!!error}
                helperText={error ?? field.help?.short}
                disabled={field.disabled || saving}
                InputLabelProps={field.type === 'date' ? { shrink: true } : undefined}
                fullWidth
              />
            );
          })}
        </Stack>
      </DialogContent>
      <DialogActions>
        <Button onClick={onCancel} disabled={saving}>
          Annulla
        </Button>
        <Button
          variant="contained"
          onClick={handleSave}
          disabled={saving}
          startIcon={saving ? <CircularProgress size={16} color="inherit" /> : undefined}
        >
          Salva
        </Button>
      </DialogActions>
    </Dialog>
  );
}
