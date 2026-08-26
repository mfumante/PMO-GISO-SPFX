import { createTheme } from '@mui/material/styles';

// Palette "Engineering": blu scuro, magenta, arancione, grigio.
const palette = {
  darkBlue: '#0B2545',
  darkBlueLight: '#13385E',
  magenta: '#A6145C',
  orange: '#E4572E',
  grey: '#5C6672',
  greyLight: '#F4F5F7',
};

export const theme = createTheme({
  palette: {
    mode: 'light',
    primary: {
      main: palette.darkBlue,
      light: palette.darkBlueLight,
      contrastText: '#FFFFFF',
    },
    secondary: {
      main: palette.magenta,
      contrastText: '#FFFFFF',
    },
    warning: {
      main: palette.orange,
      contrastText: '#FFFFFF',
    },
    background: {
      default: palette.greyLight,
      paper: '#FFFFFF',
    },
    text: {
      primary: '#1C1F26',
      secondary: palette.grey,
    },
    divider: 'rgba(11, 37, 69, 0.12)',
  },
  shape: {
    borderRadius: 8,
  },
  typography: {
    fontFamily: [
      '"Segoe UI"',
      'Roboto',
      '"Helvetica Neue"',
      'Arial',
      'sans-serif',
    ].join(','),
    h1: { fontWeight: 600 },
    h2: { fontWeight: 600 },
    h3: { fontWeight: 600 },
    h4: { fontWeight: 600 },
    h5: { fontWeight: 600 },
    h6: { fontWeight: 600 },
    button: { textTransform: 'none', fontWeight: 600 },
  },
  components: {
    MuiAppBar: {
      styleOverrides: {
        root: {
          backgroundColor: palette.darkBlue,
        },
      },
    },
    MuiDrawer: {
      styleOverrides: {
        paper: {
          backgroundColor: palette.darkBlue,
          color: '#FFFFFF',
        },
      },
    },
    MuiListItemIcon: {
      styleOverrides: {
        root: {
          color: 'rgba(255, 255, 255, 0.7)',
          minWidth: 40,
        },
      },
    },
    MuiButton: {
      styleOverrides: {
        root: {
          borderRadius: 6,
        },
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          backgroundImage: 'none',
        },
      },
    },
    // Il default di MUI (0.6875rem) e' troppo piccolo per essere letto
    // comodamente, specialmente nel popup della sidebar collassata: un solo
    // override qui rende tutti i tooltip dell'app (sidebar, azioni di
    // tabella, icone di aiuto) uniformi e leggibili, invece di sistemare la
    // dimensione caso per caso.
    MuiTooltip: {
      styleOverrides: {
        tooltip: {
          fontSize: '0.8125rem',
          padding: '6px 10px',
        },
      },
    },
  },
});
