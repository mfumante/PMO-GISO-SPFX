import * as React from 'react';
import { Box, Divider, Drawer, List, ListItemButton, ListItemIcon, ListItemText, Toolbar, Tooltip, Typography } from '@mui/material';
import { navItems, PageKey } from '../navigation/navigation';

export const DRAWER_WIDTH_EXPANDED = 240;
export const DRAWER_WIDTH_COLLAPSED = 60;

interface SidebarProps {
  currentPage: PageKey;
  onNavigate: (page: PageKey) => void;
  collapsed: boolean;
}

function SidebarContent({ currentPage, onNavigate, collapsed }: SidebarProps): React.ReactElement {
  return (
    <Box>
      <Toolbar sx={{ justifyContent: collapsed ? 'center' : 'flex-start', px: collapsed ? 1 : 2 }}>
        {!collapsed && (
          <Typography variant="h6" noWrap component="div" sx={{ fontWeight: 700 }}>
            PMO GISO
          </Typography>
        )}
      </Toolbar>
      <Divider sx={{ borderColor: 'rgba(255,255,255,0.12)' }} />
      <List>
        {navItems.map((item) => {
          const selected = currentPage === item.key;
          const Icon = item.icon;
          const button = (
            <ListItemButton
              key={item.key}
              selected={selected}
              onClick={() => onNavigate(item.key)}
              sx={{
                mx: 1,
                my: 0.5,
                borderRadius: 1,
                justifyContent: collapsed ? 'center' : 'flex-start',
                px: collapsed ? 1.5 : 2,
                color: 'rgba(255,255,255,0.85)',
                '&.Mui-selected': {
                  backgroundColor: 'secondary.main',
                  color: '#fff',
                  '& .MuiListItemIcon-root': { color: '#fff' },
                },
                '&.Mui-selected:hover': {
                  backgroundColor: 'secondary.main',
                },
                '&:hover': {
                  backgroundColor: 'rgba(255,255,255,0.08)',
                },
              }}
            >
              <ListItemIcon sx={{ minWidth: collapsed ? 0 : 40, justifyContent: 'center' }}>
                <Icon fontSize="small" />
              </ListItemIcon>
              {!collapsed && <ListItemText primary={item.label} />}
            </ListItemButton>
          );

          // Con la sidebar collassata (sole icone) il nome della voce si vede
          // solo al passaggio del mouse, tramite Tooltip.
          return collapsed ? (
            <Tooltip key={item.key} title={item.label} placement="right">
              {button}
            </Tooltip>
          ) : (
            button
          );
        })}
      </List>
    </Box>
  );
}

export function Sidebar({ currentPage, onNavigate, collapsed }: SidebarProps): React.ReactElement {
  const width = collapsed ? DRAWER_WIDTH_COLLAPSED : DRAWER_WIDTH_EXPANDED;

  return (
    <Drawer
      variant="permanent"
      sx={{
        width,
        flexShrink: 0,
        transition: (theme) =>
          theme.transitions.create('width', {
            easing: theme.transitions.easing.sharp,
            duration: theme.transitions.duration.enteringScreen,
          }),
        '& .MuiDrawer-paper': {
          // MUI ancora il Paper del Drawer al viewport ('fixed') per default:
          // lo ancoriamo invece all'antenato posizionato piu' vicino (il Box
          // radice di MainLayout), altrimenti la sidebar sfuggirebbe dal
          // riquadro della web part.
          position: 'absolute',
          boxSizing: 'border-box',
          width,
          overflowX: 'hidden',
          transition: (theme) =>
            theme.transitions.create('width', {
              easing: theme.transitions.easing.sharp,
              duration: theme.transitions.duration.enteringScreen,
            }),
        },
      }}
      open
    >
      <SidebarContent currentPage={currentPage} onNavigate={onNavigate} collapsed={collapsed} />
    </Drawer>
  );
}
