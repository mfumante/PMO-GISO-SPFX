import * as React from 'react';
import { Box, Stack, Typography } from '@mui/material';
import type { ReactNode } from 'react';

interface PageContainerProps {
  title: string;
  // Pulsanti/azioni mostrati sulla stessa riga del titolo quando c'e' spazio
  // (es. "Nuovo progetto"): evita una riga dedicata solo ai pulsanti sotto il
  // titolo, risparmiando spazio verticale.
  actions?: ReactNode;
  children?: ReactNode;
}

export function PageContainer({ title, actions, children }: PageContainerProps): React.ReactElement {
  return (
    <Box>
      <Stack
        direction="row"
        flexWrap="wrap"
        useFlexGap
        spacing={2}
        alignItems="center"
        justifyContent="space-between"
        sx={{ mb: 2 }}
      >
        <Typography variant="h5" component="h1" sx={{ fontWeight: 600 }} color="primary.main">
          {title}
        </Typography>
        {actions && <Box>{actions}</Box>}
      </Stack>
      {children}
    </Box>
  );
}
