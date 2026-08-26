import * as React from 'react';
import { Paper, Typography } from '@mui/material';

export interface IKpiCardProps {
  label: string;
  value: React.ReactNode;
  helper?: React.ReactNode;
}

// Card KPI compatta: etichetta piccola, valore in evidenza. Condivisa da
// Dashboard e Scheda Progetto (Overview) per evitare due varianti quasi
// identiche con densita' visiva diversa.
export function KpiCard({ label, value, helper }: IKpiCardProps): React.ReactElement {
  return (
    <Paper variant="outlined" sx={{ p: 1.5, height: '100%' }}>
      <Typography variant="caption" color="text.secondary" noWrap component="div">
        {label}
      </Typography>
      <Typography variant="h5" sx={{ mt: 0.25, fontWeight: 700, lineHeight: 1.2 }}>
        {value}
      </Typography>
      {helper && <Typography variant="caption" color="text.secondary">{helper}</Typography>}
    </Paper>
  );
}
