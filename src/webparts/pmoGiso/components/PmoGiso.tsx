import * as React from 'react';
import { Box } from '@mui/material';
import { StyledEngineProvider, ThemeProvider } from '@mui/material/styles';
import { theme } from './theme/theme';
import App from './App';
import type { IPmoGisoProps } from './IPmoGisoProps';

export default function PmoGiso({ context, userDisplayName }: IPmoGisoProps): React.ReactElement {
  return (
    // StyledEngineProvider injectFirst: inserisce gli stili di MUI per primi
    // nell'<head>, cosi' gli stili di Fluent UI/SharePoint della pagina host
    // restano prioritari in caso di conflitto. Niente CssBaseline globale:
    // sovrascriverebbe gli stili della pagina host.
    <StyledEngineProvider injectFirst>
      <ThemeProvider theme={theme}>
        <Box
          sx={{
            position: 'relative',
            width: '100%',
            // height: 100% permette all'albero sottostante (MainLayout) di
            // riempire tutta l'altezza disponibile quando la pagina ospitante
            // e' una App Page a piena pagina (supportsFullBleed): in quel
            // caso SharePoint da' un'altezza esplicita agli antenati del
            // riquadro della web part. In una pagina normale l'antenato ha
            // altezza 'auto', quindi questa regola non ha effetto e il
            // riquadro si dimensiona sul contenuto (vedi minHeight su
            // MainLayout per il limite minimo in quel caso).
            height: '100%',
            overflow: 'hidden',
            bgcolor: 'background.default',
          }}
        >
          <App context={context} userDisplayName={userDisplayName} />
        </Box>
      </ThemeProvider>
    </StyledEngineProvider>
  );
}
