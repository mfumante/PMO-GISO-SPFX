import { getChoiceOptions } from '../../services/SharePointDataService';
import { getFieldHelp } from '../help/fieldHelp';
import type { IFieldHelp } from '../help/fieldHelp';
import type { IFormFieldOption, IFormFieldSchema } from '../common/EntityFormDialog';

function toOptions(values: string[]): IFormFieldOption[] {
  return values.map((value) => ({ value, label: value }));
}

// Stessa spiegazione per tutti i campi persona: selezionare un utente reale del
// tenant valorizza anche il campo testo collegato; lasciare vuoto il campo
// persona e compilare solo il testo resta il modo corretto per registrare un
// soggetto esterno al tenant (consulente, fornitore).
const PERSON_SYNC_HELP: IFieldHelp = {
  short: "Cerca un utente del tenant: selezionandolo aggiorni anche il campo testo qui sopra.",
  long:
    "Usa questo campo per collegare un utente reale del tenant (necessario per notifiche e reportistica " +
    "basata su identita'). Selezionando una persona, il campo testo corrispondente viene aggiornato con il " +
    "suo nome. Per un soggetto esterno al tenant (consulente, fornitore) lascia vuoto questo campo e compila " +
    "solo il testo.",
};

// ProjectCode e' la chiave naturale referenziata da tutte le altre 5 liste:
// va bloccata in modifica per non rompere i collegamenti gia' esistenti.
export function buildProjectFields(isEditing: boolean): IFormFieldSchema[] {
  return [
    {
      name: 'ProjectCode',
      label: 'Codice progetto',
      type: 'text',
      required: true,
      disabled: isEditing,
      help: getFieldHelp('Project', 'ProjectCode'),
    },
    { name: 'Title', label: 'Titolo', type: 'text', required: true, help: getFieldHelp('Project', 'Title') },
    { name: 'Description', label: 'Descrizione', type: 'note' },
    { name: 'Sponsor', label: 'Sponsor', type: 'text', help: getFieldHelp('Project', 'Sponsor') },
    {
      name: 'SponsorUser',
      label: 'Sponsor (utente)',
      type: 'person',
      syncTextField: 'Sponsor',
      syncEmailField: 'SponsorUserEmail',
      help: PERSON_SYNC_HELP,
    },
    { name: 'SponsorUserEmail', label: 'Email Sponsor', type: 'hidden' },
    { name: 'ProjectManager', label: 'Project Manager', type: 'text', help: getFieldHelp('Project', 'ProjectManager') },
    {
      name: 'ProjectManagerUser',
      label: 'Project Manager (utente)',
      type: 'person',
      syncTextField: 'ProjectManager',
      syncEmailField: 'ProjectManagerUserEmail',
      help: PERSON_SYNC_HELP,
    },
    { name: 'ProjectManagerUserEmail', label: 'Email Project Manager', type: 'hidden' },
    { name: 'StartDate', label: 'Data inizio', type: 'date' },
    { name: 'EndDate', label: 'Data fine', type: 'date' },
    {
      name: 'Status',
      label: 'Status',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Projects', 'Status')),
      help: getFieldHelp('Project', 'Status'),
    },
    {
      name: 'Priority',
      label: 'Priority',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Projects', 'Priority')),
      help: getFieldHelp('Project', 'Priority'),
    },
    {
      name: 'RAG',
      label: 'RAG',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Projects', 'RAG')),
      help: getFieldHelp('Project', 'RAG'),
    },
    {
      name: 'Progress',
      label: 'Avanzamento (%)',
      type: 'number',
      min: 0,
      max: 100,
      help: getFieldHelp('Project', 'Progress'),
    },
    { name: 'BudgetTotal', label: 'Budget totale', type: 'currency', help: getFieldHelp('Project', 'BudgetTotal') },
    {
      name: 'BudgetCommitted',
      label: 'Budget impegnato',
      type: 'currency',
      help: getFieldHelp('Project', 'BudgetCommitted'),
    },
    {
      name: 'BudgetConsumed',
      label: 'Budget consumato',
      type: 'currency',
      help: getFieldHelp('Project', 'BudgetConsumed'),
    },
    { name: 'StrategicArea', label: 'Area strategica', type: 'text', help: getFieldHelp('Project', 'StrategicArea') },
    { name: 'Notes', label: 'Note', type: 'note' },
  ];
}

// ProjectCode e' il collegamento al progetto della Scheda da cui il form viene
// aperto: sempre precompilato e mai modificabile.
export function buildDeliverableFields(): IFormFieldSchema[] {
  return [
    { name: 'ProjectCode', label: 'Codice progetto', type: 'text', required: true, disabled: true },
    { name: 'Title', label: 'Titolo', type: 'text', required: true, help: getFieldHelp('Deliverable', 'Title') },
    { name: 'Owner', label: 'Owner', type: 'text', help: getFieldHelp('Deliverable', 'Owner') },
    {
      name: 'OwnerUser',
      label: 'Owner (utente)',
      type: 'person',
      syncTextField: 'Owner',
      syncEmailField: 'OwnerUserEmail',
      help: PERSON_SYNC_HELP,
    },
    { name: 'OwnerUserEmail', label: 'Email Owner', type: 'hidden' },
    { name: 'StartDate', label: 'Data inizio', type: 'date', help: getFieldHelp('Deliverable', 'StartDate') },
    { name: 'EndDate', label: 'Data fine', type: 'date', help: getFieldHelp('Deliverable', 'EndDate') },
    {
      name: 'Progress',
      label: 'Avanzamento (%)',
      type: 'number',
      min: 0,
      max: 100,
      help: getFieldHelp('Deliverable', 'Progress'),
    },
    {
      name: 'Status',
      label: 'Status',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Deliverables', 'Status')),
      help: getFieldHelp('Deliverable', 'Status'),
    },
    { name: 'Weight', label: 'Peso', type: 'number', help: getFieldHelp('Deliverable', 'Weight') },
    { name: 'Notes', label: 'Note', type: 'note' },
  ];
}

export function buildIssueFields(): IFormFieldSchema[] {
  return [
    { name: 'ProjectCode', label: 'Codice progetto', type: 'text', required: true, disabled: true },
    { name: 'Title', label: 'Titolo', type: 'text', required: true, help: getFieldHelp('Issue', 'Title') },
    { name: 'Description', label: 'Descrizione', type: 'note', help: getFieldHelp('Issue', 'Description') },
    {
      name: 'Severity',
      label: 'Severity',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Issues', 'Severity')),
      help: getFieldHelp('Issue', 'Severity'),
    },
    { name: 'Owner', label: 'Owner', type: 'text', help: getFieldHelp('Issue', 'Owner') },
    {
      name: 'OwnerUser',
      label: 'Owner (utente)',
      type: 'person',
      syncTextField: 'Owner',
      syncEmailField: 'OwnerUserEmail',
      help: PERSON_SYNC_HELP,
    },
    { name: 'OwnerUserEmail', label: 'Email Owner', type: 'hidden' },
    { name: 'OpenDate', label: 'Data apertura', type: 'date', help: getFieldHelp('Issue', 'OpenDate') },
    { name: 'DueDate', label: 'Scadenza', type: 'date', help: getFieldHelp('Issue', 'DueDate') },
    {
      name: 'Status',
      label: 'Status',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Issues', 'Status')),
      help: getFieldHelp('Issue', 'Status'),
    },
    { name: 'Action', label: 'Azione', type: 'note', help: getFieldHelp('Issue', 'Action') },
    {
      name: 'EscalationRequired',
      label: 'Richiede escalation',
      type: 'boolean',
      help: getFieldHelp('Issue', 'EscalationRequired'),
    },
    { name: 'Notes', label: 'Note', type: 'note' },
  ];
}

// Come in ExcelService.ts, Costs non ha una colonna Title propria nel template:
// il Title della lista (sempre presente in SharePoint) viene generato automaticamente
// da ProjectCode+CostCategory in formValuesToNewCost, non richiesto all'utente.
export function buildCostFields(): IFormFieldSchema[] {
  return [
    { name: 'ProjectCode', label: 'Codice progetto', type: 'text', required: true, disabled: true },
    {
      name: 'CostCategory',
      label: 'Categoria di costo',
      type: 'select',
      required: true,
      options: toOptions(getChoiceOptions('PMO_Costs', 'CostCategory')),
      help: getFieldHelp('Cost', 'CostCategory'),
    },
    {
      name: 'AmountAllocated',
      label: 'Importo allocato',
      type: 'currency',
      help: getFieldHelp('Cost', 'AmountAllocated'),
    },
    {
      name: 'AmountCommitted',
      label: 'Importo impegnato',
      type: 'currency',
      help: getFieldHelp('Cost', 'AmountCommitted'),
    },
    {
      name: 'AmountConsumed',
      label: 'Importo consumato',
      type: 'currency',
      help: getFieldHelp('Cost', 'AmountConsumed'),
    },
    { name: 'Supplier', label: 'Fornitore', type: 'text', help: getFieldHelp('Cost', 'Supplier') },
    { name: 'Notes', label: 'Note', type: 'note' },
  ];
}

// La risorsa va scelta da un Select popolato con PMO_Resources (mai digitata a
// mano): le opzioni sono calcolate dal chiamante, che ha gia' la lista in memoria.
// Come in ExcelService.ts, Allocations non ha una colonna Title propria nel
// template: il Title della lista viene generato da ProjectCode+ResourceCode in
// formValuesToNewAllocation, non richiesto all'utente.
export function buildAllocationFields(resourceOptions: IFormFieldOption[]): IFormFieldSchema[] {
  return [
    { name: 'ProjectCode', label: 'Codice progetto', type: 'text', required: true, disabled: true },
    {
      name: 'ResourceCode',
      label: 'Risorsa',
      type: 'select',
      required: true,
      options: resourceOptions,
      help: getFieldHelp('Allocation', 'ResourceCode'),
    },
    {
      name: 'AllocationPercent',
      label: 'Allocazione (%)',
      type: 'number',
      min: 0,
      max: 100,
      help: getFieldHelp('Allocation', 'AllocationPercent'),
    },
    {
      name: 'RoleOnProject',
      label: 'Ruolo sul progetto',
      type: 'text',
      help: getFieldHelp('Allocation', 'RoleOnProject'),
    },
    { name: 'StartDate', label: 'Data inizio', type: 'date', help: getFieldHelp('Allocation', 'StartDate') },
    { name: 'EndDate', label: 'Data fine', type: 'date', help: getFieldHelp('Allocation', 'EndDate') },
    { name: 'Notes', label: 'Note', type: 'note' },
  ];
}

// ResourceCode e' referenziato dalle Allocations: va bloccato in modifica per lo
// stesso motivo di ProjectCode su Project.
export function buildResourceFields(isEditing: boolean): IFormFieldSchema[] {
  return [
    {
      name: 'ResourceCode',
      label: 'Codice risorsa',
      type: 'text',
      required: true,
      disabled: isEditing,
      help: getFieldHelp('Resource', 'ResourceCode'),
    },
    { name: 'Title', label: 'Nome', type: 'text', required: true, help: getFieldHelp('Resource', 'Title') },
    {
      name: 'PersonUser',
      label: 'Persona (utente)',
      type: 'person',
      syncTextField: 'Title',
      syncEmailField: 'PersonUserEmail',
      help: PERSON_SYNC_HELP,
    },
    { name: 'PersonUserEmail', label: 'Email persona', type: 'hidden' },
    { name: 'Role', label: 'Ruolo', type: 'text', help: getFieldHelp('Resource', 'Role') },
    { name: 'Unit', label: 'Unita\'', type: 'text', help: getFieldHelp('Resource', 'Unit') },
    { name: 'Capacity', label: 'Capacita\'', type: 'number', help: getFieldHelp('Resource', 'Capacity') },
    { name: 'Active', label: 'Attiva', type: 'boolean', help: getFieldHelp('Resource', 'Active') },
    { name: 'Notes', label: 'Note', type: 'note' },
  ];
}
