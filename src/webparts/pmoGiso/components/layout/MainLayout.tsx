import * as React from 'react';
import { Box, Toolbar } from '@mui/material';
import type { Theme } from '@mui/material/styles';
import { useContainerWidth } from '../../hooks/useContainerWidth';
import { LayoutPreferencesProvider, useLayoutPreferences } from './LayoutPreferencesContext';
import { DRAWER_WIDTH_COLLAPSED, DRAWER_WIDTH_EXPANDED, Sidebar } from './Sidebar';
import { Topbar } from './Topbar';
import type { PageKey } from '../navigation/navigation';

interface MainLayoutProps {
  currentPage: PageKey;
  onNavigate: (page: PageKey) => void;
  userDisplayName: string;
  children?: React.ReactNode;
}

// Corpo del layout: separato da MainLayout per poter leggere sidebarCollapsed
// dal LayoutPreferencesProvider (che avvolge questo componente), da cui
// dipendono le larghezze di Topbar/Sidebar/contenuto principale.
function MainLayoutBody({ currentPage, onNavigate, userDisplayName, children }: MainLayoutProps): React.ReactElement {
  const { sidebarCollapsed } = useLayoutPreferences();
  const drawerWidth = sidebarCollapsed ? DRAWER_WIDTH_COLLAPSED : DRAWER_WIDTH_EXPANDED;
  const widthTransition = (theme: Theme): string =>
    theme.transitions.create(['width', 'margin'], {
      easing: theme.transitions.easing.sharp,
      duration: theme.transitions.duration.enteringScreen,
    });

  return (
    <>
      <Topbar userDisplayName={userDisplayName} />
      <Sidebar currentPage={currentPage} onNavigate={onNavigate} collapsed={sidebarCollapsed} />
      <Box
        component="main"
        sx={{
          flexGrow: 1,
          minWidth: 0,
          width: `calc(100% - ${drawerWidth}px)`,
          transition: widthTransition,
          bgcolor: 'background.default',
          p: { xs: 1.5, sm: 2 },
          display: 'flex',
          flexDirection: 'column',
        }}
      >
        <Toolbar />
        <Box sx={{ flexGrow: 1, minHeight: 0 }}>{children}</Box>
      </Box>
    </>
  );
}

export function MainLayout(props: MainLayoutProps): React.ReactElement {
  const { ref, width } = useContainerWidth<HTMLDivElement>();

  return (
    <Box
      ref={ref}
      sx={{
        display: 'flex',
        position: 'relative',
        width: '100%',
        // height: 100% si propaga solo se un antenato ha un'altezza esplicita
        // (come avviene in una App Page a piena pagina grazie a
        // supportsFullBleed): in un web part dentro una pagina normale
        // l'antenato ha altezza 'auto' e questa regola non ha effetto, quindi
        // il contenitore si dimensiona sul contenuto senza forzare 100vh.
        height: '100%',
        minHeight: 480,
      }}
    >
      <LayoutPreferencesProvider containerWidth={width}>
        <MainLayoutBody {...props} />
      </LayoutPreferencesProvider>
    </Box>
  );
}
