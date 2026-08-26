import * as React from 'react';
import { Stack, Tooltip, Typography } from '@mui/material';
import InfoOutlinedIcon from '@mui/icons-material/InfoOutlined';
import type { IFieldHelp } from './fieldHelp';

export interface IHelpLabelProps {
  label: React.ReactNode;
  help: IFieldHelp | undefined;
}

// Etichetta + icona informativa, riusata sia nei campi di EntityFormDialog sia
// nelle intestazioni delle tabelle. Il tooltip ha una larghezza massima cosi'
// la descrizione estesa va a capo su piu' righe invece di restare su un'unica
// riga lunghissima.
export function HelpLabel({ label, help }: IHelpLabelProps): React.ReactElement {
  if (!help) {
    return <>{label}</>;
  }
  return (
    <Stack direction="row" spacing={0.5} alignItems="center" component="span">
      <span>{label}</span>
      <Tooltip
        placement="top"
        componentsProps={{ tooltip: { sx: { maxWidth: 280 } } }}
        title={
          <Stack spacing={0.5}>
            <Typography variant="body2">{help.long}</Typography>
            {help.example && (
              <Typography variant="body2" sx={{ fontStyle: 'italic' }}>
                Esempio: {help.example}
              </Typography>
            )}
          </Stack>
        }
      >
        <InfoOutlinedIcon fontSize="inherit" sx={{ fontSize: 16, color: 'action.active', cursor: 'help' }} />
      </Tooltip>
    </Stack>
  );
}
