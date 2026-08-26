import * as React from 'react';
import { useState } from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { MainLayout } from './layout/MainLayout';
import Dashboard from './pages/Dashboard';
import PortfolioProgetti from './pages/PortfolioProgetti';
import SchedaProgetto from './pages/SchedaProgetto';
import Risorse from './pages/Risorse';
import Amministrazione from './pages/Amministrazione';
import type { PageKey } from './navigation/navigation';

export interface IAppProps {
  context: WebPartContext;
  userDisplayName: string;
}

interface IRenderPageArgs {
  page: PageKey;
  context: WebPartContext;
  selectedProjectCode: string | undefined;
  onOpenProject: (projectCode: string) => void;
  onProjectCodeChange: (projectCode: string) => void;
}

function renderPage({
  page,
  context,
  selectedProjectCode,
  onOpenProject,
  onProjectCodeChange,
}: IRenderPageArgs): React.ReactElement {
  switch (page) {
    case 'dashboard':
      return <Dashboard context={context} />;
    case 'portfolio':
      return <PortfolioProgetti context={context} onOpenProject={onOpenProject} />;
    case 'scheda-progetto':
      return (
        <SchedaProgetto
          context={context}
          projectCode={selectedProjectCode}
          onProjectCodeChange={onProjectCodeChange}
        />
      );
    case 'risorse':
      return <Risorse context={context} />;
    case 'amministrazione':
      return <Amministrazione context={context} />;
    default:
      return <Dashboard context={context} />;
  }
}

// Nessun BrowserRouter/HashRouter: la web part vive dentro una pagina
// SharePoint e non controlla l'URL del browser. La pagina attiva e' tenuta
// come stato interno e passata giu' alla Sidebar per la navigazione.
export default function App({ context, userDisplayName }: IAppProps): React.ReactElement {
  const [currentPage, setCurrentPage] = useState<PageKey>('dashboard');
  // Progetto selezionato, condiviso fra Portfolio Progetti e Scheda Progetto:
  // cliccare una riga del portfolio apre la scheda di quel progetto.
  const [selectedProjectCode, setSelectedProjectCode] = useState<string | undefined>(undefined);

  function handleOpenProject(projectCode: string): void {
    setSelectedProjectCode(projectCode);
    setCurrentPage('scheda-progetto');
  }

  return (
    <MainLayout currentPage={currentPage} onNavigate={setCurrentPage} userDisplayName={userDisplayName}>
      {renderPage({
        page: currentPage,
        context,
        selectedProjectCode,
        onOpenProject: handleOpenProject,
        onProjectCodeChange: setSelectedProjectCode,
      })}
    </MainLayout>
  );
}
