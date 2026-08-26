import { useEffect, useRef, useState } from 'react';

// Misura la larghezza reale del contenitore (non della finestra): la web part
// puo' essere stretta anche su schermi grandi (es. incassata in una colonna
// stretta di una pagina normale), quindi le soglie responsive delle tabelle e
// dei grafici devono reagire a questa larghezza, non a quella del browser.
export function useContainerWidth<T extends HTMLElement>(): { ref: React.RefObject<T>; width: number } {
  const ref = useRef<T>(null);
  const [width, setWidth] = useState(0);

  useEffect(() => {
    const element = ref.current;
    if (!element) {
      return undefined;
    }

    function measure(): void {
      if (element) {
        setWidth(element.clientWidth);
      }
    }

    measure();

    // ResizeObserver non e' garantito in ogni ambiente host: se assente si
    // ricade sul resize della finestra, che copre comunque il caso piu'
    // comune (ridimensionamento del browser) anche se non intercetta
    // variazioni di larghezza dovute solo al layout circostante.
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(() => measure());
      observer.observe(element);
      return () => observer.disconnect();
    }

    window.addEventListener('resize', measure);
    return () => window.removeEventListener('resize', measure);
  }, []);

  return { ref, width };
}
