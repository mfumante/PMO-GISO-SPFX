// Guida contestuale ai campi, usata da EntityFormDialog (helperText + tooltip),
// dalle intestazioni delle tabelle e dalla card "Guida alla compilazione" in
// Amministrazione. Un'unica fonte di testo per tutti e tre i punti di utilizzo,
// cosi' la descrizione di un campo non deve essere aggiornata in piu' posti.

export type HelpEntity = 'Project' | 'Deliverable' | 'Issue' | 'Resource' | 'Allocation' | 'Cost';

export interface IFieldHelp {
  // Una riga, mostrata come helperText sotto il campo nel form.
  short: string;
  // Descrizione estesa mostrata nel tooltip dell'icona informativa.
  long: string;
  // Esempio concreto di compilazione, dove utile.
  example?: string;
}

type EntityFieldHelp = Record<string, IFieldHelp>;

const PROJECT_HELP: EntityFieldHelp = {
  ProjectCode: {
    short: 'Codice univoco del progetto, non modificabile dopo la creazione.',
    long:
      'Identifica il progetto in modo univoco ed e\' la chiave usata per collegare deliverable, issue, ' +
      'allocazioni e costi. Per questo non puo\' essere cambiato dopo la creazione.',
    example: 'PRJ001',
  },
  Title: {
    short: 'Denominazione estesa del progetto.',
    long: 'Il nome con cui il progetto viene identificato nel portfolio e nella reportistica.',
    example: "Monitoraggio NIS2 Societa' del Gruppo",
  },
  Sponsor: {
    short: 'Figura che ha commissionato il progetto e ne risponde a livello di governance.',
    long:
      "Lo sponsor rappresenta il progetto nei confronti della governance aziendale ed e' il riferimento " +
      'per le decisioni strategiche che lo riguardano.',
  },
  ProjectManager: {
    short: "Responsabile operativo dell'esecuzione e del monitoraggio.",
    long:
      'Il project manager segue l\'avanzamento quotidiano del progetto, coordina le risorse assegnate e ' +
      "aggiorna lo stato di deliverable, issue e budget.",
  },
  Status: {
    short: 'Stato di avanzamento complessivo del progetto.',
    long: 'Riflette la fase del ciclo di vita in cui si trova il progetto, indipendentemente dal RAG.',
  },
  Priority: {
    short: "Rilevanza del progetto rispetto agli obiettivi della unit.",
    long: "Aiuta a stabilire le priorita' quando risorse e budget sono limitati.",
  },
  RAG: {
    short: 'Semaforo di salute del progetto (Green/Amber/Red/Grey).',
    long:
      'Green = in linea con piani e budget. Amber = criticita\' gestibili, richiede attenzione. ' +
      "Red = criticita' che compromettono tempi, costi o obiettivi. Grey = non ancora avviato o non valutabile.",
  },
  Progress: {
    short: 'Percentuale di avanzamento complessiva, da 0 a 100.',
    long: 'Valore stimato sulla base dei deliverable completati.',
  },
  BudgetTotal: {
    short: 'Budget complessivo allocato al progetto.',
    long: "L'importo totale approvato per l'intero ciclo di vita del progetto.",
  },
  BudgetCommitted: {
    short: 'Quota gia\' impegnata contrattualmente, anche se non ancora fatturata.',
    long:
      'Include ordini e contratti firmati che generano un impegno di spesa, anche se la fattura non e\' ' +
      'ancora stata emessa o pagata.',
  },
  BudgetConsumed: {
    short: 'Quota effettivamente consumata e rendicontata.',
    long: "La parte di budget gia' spesa e documentata, ad esempio con fatture ricevute.",
  },
  StrategicArea: {
    short: 'Area di riferimento del progetto.',
    long: "Raggruppa i progetti per ambito, utile per le analisi di portfolio.",
    example: 'Compliance, Sicurezza tecnica, Governance',
  },
};

const DELIVERABLE_HELP: EntityFieldHelp = {
  Title: {
    short: 'Nome del deliverable o della milestone.',
    long: "Descrive sinteticamente cosa verra' prodotto o raggiunto.",
  },
  Owner: {
    short: 'Responsabile della produzione del deliverable.',
    long: "La persona a cui e' affidata la realizzazione di questo specifico deliverable.",
  },
  StartDate: {
    short: 'Data pianificata di inizio.',
    long: 'La data in cui e\' previsto iniziare la lavorazione del deliverable.',
  },
  EndDate: {
    short: 'Data pianificata di fine.',
    long: 'La data entro cui il deliverable dovrebbe essere completato.',
  },
  Progress: {
    short: 'Avanzamento del singolo deliverable, da 0 a 100.',
    long: 'Indica quanto lavoro e\' stato completato su questo specifico deliverable.',
  },
  Status: {
    short: 'Stato del deliverable.',
    long: 'Delayed indica uno scostamento rispetto alla data pianificata di fine.',
  },
  Weight: {
    short: 'Peso relativo del deliverable sul progetto.',
    long:
      'Utile quando i deliverable non hanno pari rilevanza sull\'avanzamento complessivo del progetto. ' +
      'Se non valorizzato si assume un peso uniforme tra tutti i deliverable.',
  },
};

const ISSUE_HELP: EntityFieldHelp = {
  Title: {
    short: "Sintesi breve della criticita'.",
    long: "Il titolo con cui la criticita' viene identificata negli elenchi e nei report.",
  },
  Description: {
    short: 'Descrizione estesa del problema e del contesto.',
    long: "Riporta cosa e' successo, l'impatto osservato e ogni informazione utile a comprendere la criticita'.",
  },
  Severity: {
    short: "Impatto della criticita' sul progetto.",
    long:
      'Critical = blocca il progetto. High = impatto rilevante su tempi o costi. ' +
      'Medium = impatto contenuto e gestibile. Low = impatto marginale.',
  },
  Owner: {
    short: 'Responsabile della risoluzione.',
    long: "La persona incaricata di seguire e chiudere la criticita'.",
  },
  OpenDate: {
    short: 'Data di rilevazione.',
    long: "La data in cui la criticita' e' stata identificata e registrata.",
  },
  DueDate: {
    short: "Data entro cui la criticita' deve essere risolta.",
    long: 'La scadenza concordata per la risoluzione o la mitigazione.',
  },
  Status: {
    short: 'Stato di gestione.',
    long: 'Mitigated indica rischio ridotto ma non eliminato.',
  },
  Action: {
    short: 'Azione correttiva definita.',
    long: "Descrive cosa e' stato deciso per risolvere o mitigare la criticita'.",
  },
  EscalationRequired: {
    short: "Indica se la criticita' richiede escalation al livello superiore.",
    long:
      'Da attivare quando la criticita\' non puo\' essere gestita a livello di progetto e serve il ' +
      'coinvolgimento della governance superiore.',
  },
};

const RESOURCE_HELP: EntityFieldHelp = {
  ResourceCode: {
    short: 'Identificativo univoco della risorsa, non modificabile dopo la creazione.',
    long:
      "E' la chiave usata per collegare le allocazioni a questa risorsa: per questo non puo' essere " +
      'cambiata dopo la creazione.',
    example: 'RIS001',
  },
  Title: {
    short: 'Nome e cognome della risorsa.',
    long: 'Il nominativo con cui la risorsa viene identificata negli elenchi e nelle allocazioni.',
  },
  Role: {
    short: 'Ruolo ricoperto.',
    long: "Il ruolo abituale della risorsa, indipendente dal ruolo che puo' ricoprire su un singolo progetto.",
  },
  Unit: {
    short: "Unita' organizzativa di appartenenza.",
    long: "La struttura aziendale a cui la risorsa afferisce.",
  },
  Capacity: {
    short: 'Capacita\' massima allocabile, in percentuale.',
    long: 'Normalmente 100. Un valore inferiore indica una disponibilita\' parziale (es. part-time).',
  },
  Active: {
    short: 'Indica se la risorsa e\' attualmente disponibile per nuove allocazioni.',
    long: "Le risorse non attive restano visibili per lo storico ma non dovrebbero ricevere nuove allocazioni.",
  },
};

const ALLOCATION_HELP: EntityFieldHelp = {
  ResourceCode: {
    short: 'Risorsa allocata sul progetto.',
    long: 'La risorsa selezionata da PMO_Resources che viene impegnata su questo progetto.',
  },
  AllocationPercent: {
    short: 'Percentuale di impegno della risorsa sul progetto.',
    long: "La somma delle allocazioni di una risorsa su tutti i progetti non dovrebbe superare la sua Capacity.",
  },
  RoleOnProject: {
    short: 'Ruolo specifico ricoperto dalla risorsa su questo progetto.',
    long: "Puo' differire dal ruolo abituale della risorsa registrato in PMO_Resources.",
  },
  StartDate: {
    short: 'Inizio del periodo di validita\' dell\'allocazione.',
    long: "La data da cui l'allocazione della risorsa sul progetto e' considerata attiva.",
  },
  EndDate: {
    short: 'Fine del periodo di validita\' dell\'allocazione.',
    long: "La data fino a cui l'allocazione della risorsa sul progetto e' considerata attiva.",
  },
};

const COST_HELP: EntityFieldHelp = {
  CostCategory: {
    short: 'Natura della voce di costo.',
    long: "Classifica il costo per tipologia, utile per l'analisi dei costi tra progetti.",
  },
  AmountAllocated: {
    short: 'Importo previsto a budget per la categoria.',
    long: "L'importo stanziato per questa categoria di costo all'interno del budget di progetto.",
  },
  AmountCommitted: {
    short: 'Importo gia\' impegnato contrattualmente.',
    long: 'Include ordini e contratti firmati relativi a questa categoria di costo.',
  },
  AmountConsumed: {
    short: 'Importo effettivamente consumato.',
    long: "La parte di questa categoria di costo gia' spesa e documentata.",
  },
  Supplier: {
    short: 'Fornitore o societa\' di riferimento.',
    long: 'Il fornitore o la societa\' di consulenza associata a questa voce di costo, se applicabile.',
    example: 'EY, Protiviti, Balance',
  },
};

export const FIELD_HELP: Record<HelpEntity, EntityFieldHelp> = {
  Project: PROJECT_HELP,
  Deliverable: DELIVERABLE_HELP,
  Issue: ISSUE_HELP,
  Resource: RESOURCE_HELP,
  Allocation: ALLOCATION_HELP,
  Cost: COST_HELP,
};

export function getFieldHelp(entity: HelpEntity, fieldName: string): IFieldHelp | undefined {
  return FIELD_HELP[entity][fieldName];
}

// Ordine ed etichette per la card "Guida alla compilazione" in Amministrazione.
export const HELP_ENTITIES: { key: HelpEntity; label: string }[] = [
  { key: 'Project', label: 'Progetti' },
  { key: 'Deliverable', label: 'Deliverable' },
  { key: 'Issue', label: 'Issue' },
  { key: 'Resource', label: 'Risorse' },
  { key: 'Allocation', label: 'Allocazioni' },
  { key: 'Cost', label: 'Costi' },
];
