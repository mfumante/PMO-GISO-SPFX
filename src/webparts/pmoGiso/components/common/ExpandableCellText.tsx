import * as React from 'react';
import { useState } from 'react';
import { ClickAwayListener, Tooltip, Typography } from '@mui/material';

export interface IExpandableCellTextProps {
  text: string | undefined;
  fallback?: string;
  // Ferma la propagazione del click oltre al testo: serve nelle righe di
  // tabella che hanno anche un onClick proprio (es. apertura del progetto in
  // Portfolio), cosi' cliccare il testo mostra il tooltip invece di innescare
  // anche l'azione della riga.
  stopPropagation?: boolean;
}

// Testo di una cella di tabella troncato con ellissi (vedi tableStyles.ts):
// un click apre un tooltip con il testo completo. A differenza del solo
// attributo title nativo (richiede hover prolungato e non funziona su touch),
// questo resta utilizzabile anche su schermi tattili e da tastiera.
export function ExpandableCellText({ text, fallback, stopPropagation }: IExpandableCellTextProps): React.ReactElement {
  const [open, setOpen] = useState(false);

  if (!text) {
    return <>{fallback ?? '-'}</>;
  }

  return (
    <ClickAwayListener onClickAway={() => setOpen(false)}>
      <Tooltip
        title={text}
        open={open}
        onClose={() => setOpen(false)}
        disableFocusListener
        disableHoverListener
        disableTouchListener
        placement="top-start"
        arrow
      >
        <Typography
          variant="body2"
          component="span"
          onClick={(event) => {
            if (stopPropagation) {
              event.stopPropagation();
            }
            setOpen((prev) => !prev);
          }}
          sx={{
            display: 'block',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            cursor: 'pointer',
          }}
        >
          {text}
        </Typography>
      </Tooltip>
    </ClickAwayListener>
  );
}
