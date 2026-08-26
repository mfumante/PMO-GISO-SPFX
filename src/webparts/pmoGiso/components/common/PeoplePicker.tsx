import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  Avatar,
  Chip,
  CircularProgress,
  ClickAwayListener,
  List,
  ListItemButton,
  Paper,
  Stack,
  TextField,
  Typography,
} from '@mui/material';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { PeopleService } from '../../services/PeopleService';
import type { IPersonSearchResult } from '../../services/PeopleService';

export interface IPersonValue {
  // Noto quando il valore proviene da un record gia' salvato: in questo caso
  // l'Id e' gia' risolto e non serve richiamare ensureUser al salvataggio.
  id?: number;
  displayName: string;
  email: string;
  // Noto quando il valore proviene da una ricerca people picker appena fatta:
  // serve a risolvere l'Id numerico tramite ensureUser al salvataggio.
  loginName?: string;
}

export interface IPeoplePickerProps {
  context: WebPartContext;
  label: string;
  value: IPersonValue | undefined;
  onChange: (value: IPersonValue | undefined) => void;
  disabled?: boolean;
  error?: boolean;
  helperText?: string;
}

const DEBOUNCE_MS = 350;
const MIN_QUERY_LENGTH = 2;

export function getInitials(name: string): string {
  const parts = name.trim().split(' ').filter((part) => part.length > 0);
  if (parts.length === 0) {
    return '?';
  }
  const first = parts[0].charAt(0);
  const last = parts.length > 1 ? parts[parts.length - 1].charAt(0) : '';
  return (first + last).toUpperCase();
}

// People picker senza Autocomplete/Popper di MUI: in alcuni contesti di
// hosting SPFx l'input dell'Autocomplete risultava non digitabile (nessun
// errore JS, in edit e in anteprima), molto probabilmente per un conflitto
// fra il Portal del popup e la pagina ospitante. Qui il campo di ricerca e'
// un TextField 'nudo' e l'elenco dei risultati e' un pannello posizionato in
// assoluto ma ancorato allo stesso contenitore (nessun Portal), sotto forma
// di lista cliccabile. Si tiene sempre un solo valore selezionato (mostrato
// come Chip rimovibile sotto il campo), sostituito da una nuova selezione.
export function PeoplePicker({
  context,
  label,
  value,
  onChange,
  disabled,
  error,
  helperText,
}: IPeoplePickerProps): React.ReactElement {
  const service = useMemo(() => new PeopleService(context), [context]);
  const [inputValue, setInputValue] = useState('');
  const [options, setOptions] = useState<IPersonSearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [open, setOpen] = useState(false);
  const debounceRef = useRef<number | undefined>(undefined);

  useEffect(() => {
    const trimmed = inputValue.trim();

    if (debounceRef.current != null) {
      window.clearTimeout(debounceRef.current);
      debounceRef.current = undefined;
    }

    if (trimmed.length < MIN_QUERY_LENGTH) {
      setOptions([]);
      setLoading(false);
      return;
    }

    setLoading(true);
    debounceRef.current = window.setTimeout(() => {
      // Niente Promise.prototype.finally: la lib TS di questo progetto esclude
      // es2018.promise. setLoading(false) va ripetuto sia in .then() sia in
      // .catch(), altrimenti in caso di errore lo spinner resterebbe attivo.
      service
        .searchUsers(trimmed)
        .then((results) => {
          setOptions(results);
          setLoading(false);
        })
        .catch(() => {
          setOptions([]);
          setLoading(false);
        });
    }, DEBOUNCE_MS);

    return () => {
      if (debounceRef.current != null) {
        window.clearTimeout(debounceRef.current);
      }
    };
  }, [inputValue, service]);

  function handleSelect(option: IPersonSearchResult): void {
    onChange({ displayName: option.displayName, email: option.email, loginName: option.loginName });
    setInputValue('');
    setOptions([]);
    setOpen(false);
  }

  function handleRemove(): void {
    onChange(undefined);
  }

  const trimmedInput = inputValue.trim();

  return (
    <ClickAwayListener onClickAway={() => setOpen(false)}>
      <Stack spacing={0.5} sx={{ position: 'relative' }}>
        <TextField
          label={label}
          size="small"
          fullWidth
          value={inputValue}
          disabled={disabled}
          error={error}
          helperText={helperText}
          onFocus={() => setOpen(true)}
          onChange={(event) => {
            setInputValue(event.target.value);
            setOpen(true);
          }}
          InputProps={{
            endAdornment: loading ? <CircularProgress color="inherit" size={16} /> : undefined,
          }}
        />
        {value && (
          <Stack direction="row">
            <Chip
              size="small"
              avatar={<Avatar>{getInitials(value.displayName)}</Avatar>}
              label={value.displayName}
              onDelete={disabled ? undefined : handleRemove}
            />
          </Stack>
        )}
        {open && (
          <Paper
            elevation={4}
            sx={{
              position: 'absolute',
              top: '100%',
              left: 0,
              right: 0,
              zIndex: 10,
              mt: 0.5,
              maxHeight: 240,
              overflowY: 'auto',
            }}
          >
            {trimmedInput.length < MIN_QUERY_LENGTH ? (
              <Typography variant="body2" color="text.secondary" sx={{ px: 2, py: 1 }}>
                Digita per cercare
              </Typography>
            ) : options.length === 0 ? (
              <Typography variant="body2" color="text.secondary" sx={{ px: 2, py: 1 }}>
                {loading ? 'Ricerca in corso...' : 'Nessun utente trovato'}
              </Typography>
            ) : (
              <List dense disablePadding>
                {options.map((option) => (
                  <ListItemButton key={option.loginName} onClick={() => handleSelect(option)}>
                    <Stack direction="row" spacing={1} alignItems="center">
                      <Avatar sx={{ width: 28, height: 28, fontSize: 13 }}>{getInitials(option.displayName)}</Avatar>
                      <Stack spacing={0}>
                        <Typography variant="body2">{option.displayName}</Typography>
                        {option.email && (
                          <Typography variant="caption" color="text.secondary">
                            {option.email}
                          </Typography>
                        )}
                      </Stack>
                    </Stack>
                  </ListItemButton>
                ))}
              </List>
            )}
          </Paper>
        )}
      </Stack>
    </ClickAwayListener>
  );
}
