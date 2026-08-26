import * as React from 'react';
import MenuIcon from '@mui/icons-material/Menu';
import MenuOpenIcon from '@mui/icons-material/MenuOpen';
import { AppBar, Box, IconButton, ToggleButton, ToggleButtonGroup, Toolbar, Tooltip, Typography } from '@mui/material';
import type { MouseEvent } from 'react';
import { DRAWER_WIDTH_COLLAPSED, DRAWER_WIDTH_EXPANDED } from './Sidebar';
import { useLayoutPreferences } from './LayoutPreferencesContext';
import type { TableDensity } from './LayoutPreferencesContext';

interface TopbarProps {
  userDisplayName: string;
}

export function Topbar({ userDisplayName }: TopbarProps): React.ReactElement {
  const { sidebarCollapsed, setSidebarCollapsed, density, setDensity } = useLayoutPreferences();
  const drawerWidth = sidebarCollapsed ? DRAWER_WIDTH_COLLAPSED : DRAWER_WIDTH_EXPANDED;

  function handleDensityChange(_event: MouseEvent<HTMLElement>, value: TableDensity | null): void {
    if (value) {
      setDensity(value);
    }
  }

  return (
    <AppBar
      // 'fixed' ancorerebbe la barra al viewport del browser invece che al
      // riquadro della web part: 'absolute' la ancora al contenitore relativo
      // piu' vicino (MainLayout).
      position="absolute"
      sx={{
        width: `calc(100% - ${drawerWidth}px)`,
        ml: `${drawerWidth}px`,
        transition: (theme) =>
          theme.transitions.create(['width', 'margin'], {
            easing: theme.transitions.easing.sharp,
            duration: theme.transitions.duration.enteringScreen,
          }),
      }}
    >
      <Toolbar sx={{ gap: { xs: 1, sm: 2 } }}>
        <Tooltip title={sidebarCollapsed ? 'Espandi il menu' : 'Comprimi il menu'}>
          <IconButton color="inherit" edge="start" onClick={() => setSidebarCollapsed(!sidebarCollapsed)}>
            {sidebarCollapsed ? <MenuIcon /> : <MenuOpenIcon />}
          </IconButton>
        </Tooltip>
        <Typography variant="h6" noWrap component="div" sx={{ flexGrow: 1, display: { xs: 'none', sm: 'block' } }}>
          Gestione Progetti GISO
        </Typography>
        <ToggleButtonGroup
          size="small"
          value={density}
          exclusive
          onChange={handleDensityChange}
          sx={{
            bgcolor: 'rgba(255,255,255,0.08)',
            '& .MuiToggleButton-root': { color: 'inherit', textTransform: 'none', px: 1.5, border: 0 },
            '& .Mui-selected': { bgcolor: 'rgba(255,255,255,0.28) !important', fontWeight: 600 },
          }}
        >
          <ToggleButton value="comfortable">Comoda</ToggleButton>
          <ToggleButton value="compact">Compatta</ToggleButton>
        </ToggleButtonGroup>
        <Box sx={{ display: { xs: 'none', md: 'block' } }}>
          <Typography variant="body2" noWrap sx={{ opacity: 0.85 }}>
            {userDisplayName}
          </Typography>
        </Box>
      </Toolbar>
    </AppBar>
  );
}
