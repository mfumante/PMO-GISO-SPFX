import type { SxProps, Theme } from '@mui/material';
import type { TableDensity } from '../layout/LayoutPreferencesContext';

// Altezza massima dell'area scrollabile delle tabelle: proporzionale al
// viewport (non un pixel fisso) cosi' l'intestazione sticky ha davvero una
// finestra di scroll verticale in cui restare ancorata, sia in un web part
// incassato sia in una App Page a piena altezza.
export const TABLE_MAX_HEIGHT = '60vh';

// Stile condiviso da tutte le tabelle dell'app (Portfolio, Scheda Progetto,
// Risorse, Dashboard): intestazione sticky durante lo scroll verticale, righe
// alternate per leggibilita', testo troncato con ellissi (richiede
// table-layout fixed sulla <Table>, applicato qui) e padding proporzionato
// alla densita' scelta dall'utente nella topbar.
//
// minWidth: senza un minimo esplicito la <Table> resta sempre larga il 100%
// del contenitore (le percentuali di colonna si dividono quello spazio),
// quindi non trabocca mai e lo scroll orizzontale di TableContainer non si
// attiva mai. Impostando un minWidth (somma delle larghezze minime ragionevoli
// delle colonne visibili) la tabella smette di restringersi oltre quel punto:
// su contenitori piu' stretti trabocca e diventa scorrevole lateralmente,
// invece di schiacciare il contenuto illeggibile.
export function getTableSx(density: TableDensity, minWidth?: number): SxProps<Theme> {
  const cellPaddingY = density === 'compact' ? '4px' : '8px';
  return {
    tableLayout: 'fixed',
    minWidth: minWidth ?? undefined,
    '& .MuiTableCell-root': {
      paddingTop: cellPaddingY,
      paddingBottom: cellPaddingY,
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      whiteSpace: 'nowrap',
    },
    '& thead .MuiTableCell-root': {
      position: 'sticky',
      top: 0,
      zIndex: 1,
      backgroundColor: 'background.paper',
      fontWeight: 600,
      whiteSpace: 'nowrap',
    },
    '& tbody .MuiTableRow-root:nth-of-type(odd)': {
      backgroundColor: 'action.hover',
    },
  };
}
