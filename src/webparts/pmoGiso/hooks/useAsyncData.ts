import { useEffect, useState } from 'react';
import type { DependencyList } from 'react';
import { SharePointListNotFoundError } from '../services/SharePointDataService';

export type AsyncDataState<T> =
  | { status: 'loading' }
  | { status: 'error'; message: string }
  | { status: 'missing'; listTitle: string }
  | { status: 'ready'; data: T };

// Centralizza il pattern di caricamento condiviso da tutte le pagine dati: stato di
// caricamento, distinzione fra "lista non trovata" (404, da segnalare senza crash)
// ed errore generico, e annullamento del risultato se il componente si smonta prima
// che la Promise si risolva.
export function useAsyncData<T>(loader: () => Promise<T>, deps: DependencyList): AsyncDataState<T> {
  const [state, setState] = useState<AsyncDataState<T>>({ status: 'loading' });

  useEffect(() => {
    let cancelled = false;
    setState({ status: 'loading' });

    loader()
      .then((data) => {
        if (!cancelled) {
          setState({ status: 'ready', data });
        }
      })
      .catch((error: unknown) => {
        if (cancelled) {
          return;
        }
        if (error instanceof SharePointListNotFoundError) {
          setState({ status: 'missing', listTitle: error.listTitle });
        } else {
          setState({ status: 'error', message: error instanceof Error ? error.message : 'Errore imprevisto.' });
        }
      });

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  return state;
}
