import DashboardOutlinedIcon from '@mui/icons-material/DashboardOutlined';
import AccountTreeOutlinedIcon from '@mui/icons-material/AccountTreeOutlined';
import AssignmentOutlinedIcon from '@mui/icons-material/AssignmentOutlined';
import GroupsOutlinedIcon from '@mui/icons-material/GroupsOutlined';
import AdminPanelSettingsOutlinedIcon from '@mui/icons-material/AdminPanelSettingsOutlined';
import type SvgIcon from '@mui/material/SvgIcon';

// Chiave di pagina interna: la web part non controlla l'URL della pagina host,
// quindi la navigazione avviene tramite stato React invece che tramite rotte.
export type PageKey = 'dashboard' | 'portfolio' | 'scheda-progetto' | 'risorse' | 'amministrazione';

export interface NavItem {
  key: PageKey;
  label: string;
  // '@mui/icons-material' non esporta pubblicamente il tipo dei suoi componenti
  // icona: lo si ricava da SvgIcon di @mui/material, di cui condividono la firma.
  icon: typeof SvgIcon;
}

export const navItems: NavItem[] = [
  { key: 'dashboard', label: 'Dashboard', icon: DashboardOutlinedIcon },
  { key: 'portfolio', label: 'Portfolio Progetti', icon: AccountTreeOutlinedIcon },
  { key: 'scheda-progetto', label: 'Scheda Progetto', icon: AssignmentOutlinedIcon },
  { key: 'risorse', label: 'Risorse', icon: GroupsOutlinedIcon },
  { key: 'amministrazione', label: 'Amministrazione', icon: AdminPanelSettingsOutlinedIcon },
];
