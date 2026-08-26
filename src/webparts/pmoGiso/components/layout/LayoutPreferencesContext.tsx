import * as React from 'react';
import { createContext, useContext, useEffect, useState } from 'react';

export type WidthTier = 'narrow' | 'medium' | 'wide';
export type TableDensity = 'comfortable' | 'compact';

// Soglie di larghezza del CONTENITORE (non della finestra): sotto NARROW_MAX
// fascia "stretta", fino a MEDIUM_MAX fascia "media", oltre fascia "ampia".
const NARROW_MAX = 800;
const MEDIUM_MAX = 1300;

export function widthToTier(width: number): WidthTier {
  if (width > 0 && width < NARROW_MAX) {
    return 'narrow';
  }
  if (width <= MEDIUM_MAX) {
    return 'medium';
  }
  return 'wide';
}

const SIDEBAR_STORAGE_KEY = 'pmoGiso.sidebarCollapsed';
const DENSITY_STORAGE_KEY = 'pmoGiso.tableDensity';

// sessionStorage puo' non essere disponibile (modalita' privata, policy del
// browser/host): in quel caso le preferenze semplicemente non persistono tra
// un refresh e l'altro, l'app resta comunque pienamente utilizzabile.
function readSessionFlag(key: string): boolean | undefined {
  try {
    const raw = window.sessionStorage.getItem(key);
    return raw == null ? undefined : raw === '1';
  } catch {
    return undefined;
  }
}

function writeSessionFlag(key: string, value: boolean): void {
  try {
    window.sessionStorage.setItem(key, value ? '1' : '0');
  } catch {
    // Vedi commento sopra.
  }
}

function readSessionDensity(): TableDensity | undefined {
  try {
    const raw = window.sessionStorage.getItem(DENSITY_STORAGE_KEY);
    return raw === 'compact' || raw === 'comfortable' ? raw : undefined;
  } catch {
    return undefined;
  }
}

function writeSessionDensity(value: TableDensity): void {
  try {
    window.sessionStorage.setItem(DENSITY_STORAGE_KEY, value);
  } catch {
    // Vedi commento sopra readSessionFlag.
  }
}

export interface ILayoutPreferences {
  widthTier: WidthTier;
  containerWidth: number;
  sidebarCollapsed: boolean;
  setSidebarCollapsed: (collapsed: boolean) => void;
  density: TableDensity;
  setDensity: (density: TableDensity) => void;
}

const LayoutPreferencesContext = createContext<ILayoutPreferences | undefined>(undefined);

interface LayoutPreferencesProviderProps {
  containerWidth: number;
  children: React.ReactNode;
}

// Fornisce a tutta la web part (Topbar, Sidebar, pagine) lo stato di layout
// condiviso, evitando di far attraversare queste preferenze da MainLayout fino
// alle pagine annidate tramite prop drilling.
export function LayoutPreferencesProvider({
  containerWidth,
  children,
}: LayoutPreferencesProviderProps): React.ReactElement {
  const widthTier = widthToTier(containerWidth);

  const [sidebarCollapsed, setSidebarCollapsedState] = useState<boolean>(
    () => readSessionFlag(SIDEBAR_STORAGE_KEY) ?? false,
  );
  const [density, setDensityState] = useState<TableDensity>(() => readSessionDensity() ?? 'comfortable');

  // Su larghezze contenute la sidebar deve partire gia' collassata. Si applica
  // solo finche' l'utente non ha espresso esplicitamente una preferenza in
  // questa sessione, per non sovrascrivere una scelta manuale ad ogni
  // ridimensionamento del contenitore.
  const [hasUserToggled, setHasUserToggled] = useState<boolean>(() => readSessionFlag(SIDEBAR_STORAGE_KEY) != null);

  useEffect(() => {
    if (!hasUserToggled && widthTier === 'narrow') {
      setSidebarCollapsedState(true);
    }
  }, [widthTier, hasUserToggled]);

  function setSidebarCollapsed(collapsed: boolean): void {
    setSidebarCollapsedState(collapsed);
    setHasUserToggled(true);
    writeSessionFlag(SIDEBAR_STORAGE_KEY, collapsed);
  }

  function setDensity(value: TableDensity): void {
    setDensityState(value);
    writeSessionDensity(value);
  }

  const contextValue: ILayoutPreferences = {
    widthTier,
    containerWidth,
    sidebarCollapsed,
    setSidebarCollapsed,
    density,
    setDensity,
  };

  return <LayoutPreferencesContext.Provider value={contextValue}>{children}</LayoutPreferencesContext.Provider>;
}

export function useLayoutPreferences(): ILayoutPreferences {
  const ctx = useContext(LayoutPreferencesContext);
  if (!ctx) {
    throw new Error('useLayoutPreferences deve essere usato dentro LayoutPreferencesProvider.');
  }
  return ctx;
}
