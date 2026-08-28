import * as XLSX from 'xlsx';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { PeopleService } from './PeopleService';
import {
  CostCategory,
  DeliverableStatus,
  IAllocation,
  ICost,
  IDeliverable,
  IIssue,
  IProject,
  IResource,
  IssueSeverity,
  IssueStatus,
  IWriteResult,
  NewAllocation,
  NewCost,
  NewDeliverable,
  NewIssue,
  NewProject,
  NewResource,
  ProjectPriority,
  ProjectRag,
  ProjectStatus,
  SharePointDataService,
  getChoiceOptions,
} from './SharePointDataService';

// ---------------------------------------------------------------------------
// Tipi pubblici
// ---------------------------------------------------------------------------

export type SheetName = 'Projects' | 'Deliverables' | 'Issues' | 'Resources' | 'Allocations' | 'Costs';

export interface IParsedWorkbook {
  sheets: Record<SheetName, Record<string, unknown>[]>;
  missingSheets: SheetName[];
}

export type ValidationLevel = 'error' | 'warning';

export interface IValidationIssue {
  sheet: SheetName;
  row: number;
  column?: string;
  level: ValidationLevel;
  message: string;
}

export interface IValidatedRow<T> {
  rowNumber: number;
  isUpdate: boolean;
  data: T;
}

export interface ISheetValidationResult<T> {
  totalRows: number;
  validRows: IValidatedRow<T>[];
  errors: IValidationIssue[];
  warnings: IValidationIssue[];
}

export interface IValidationReport {
  Projects: ISheetValidationResult<NewProject>;
  Deliverables: ISheetValidationResult<NewDeliverable>;
  Issues: ISheetValidationResult<NewIssue>;
  Resources: ISheetValidationResult<NewResource>;
  Allocations: ISheetValidationResult<NewAllocation>;
  Costs: ISheetValidationResult<NewCost>;
  hasBlockingErrors: boolean;
  missingSheets: SheetName[];
}

export interface IImportLogEntry {
  sheet: SheetName;
  row?: number;
  level: 'info' | 'success' | 'error';
  message: string;
}

export interface ISheetImportResult {
  created: number;
  updated: number;
  skipped: number;
  errors: IValidationIssue[];
}

export interface IImportReport {
  Projects: ISheetImportResult;
  Deliverables: ISheetImportResult;
  Issues: ISheetImportResult;
  Resources: ISheetImportResult;
  Allocations: ISheetImportResult;
  Costs: ISheetImportResult;
}

export interface IImportCallbacks {
  onLog?: (entry: IImportLogEntry) => void;
  onProgress?: (completed: number, total: number) => void;
}

// Lanciato quando il file selezionato non e' un file Excel leggibile.
export class ExcelParseError extends Error {}

// Alias derivato via ReturnType invece di riferire XLSX.WorkBook per nome: il
// pacchetto xlsx non e' installato in questo ambiente di sviluppo (vedi nota nel
// riepilogo), quindi non e' stato possibile verificarne i tipi esatti.
type Workbook = ReturnType<typeof XLSX.utils.book_new>;

// ---------------------------------------------------------------------------
// Definizione colonne (stessa struttura per template, parsing ed export)
// ---------------------------------------------------------------------------

// Le colonne *Email sono opzionali e affiancano i campi testo esistenti (non li
// sostituiscono): se valorizzate, in import vengono risolte in un utente reale
// del tenant tramite PeopleService (vedi validateProjectsSheet e affini).
const PROJECT_COLUMNS = [
  'ProjectCode',
  'Title',
  'Description',
  'Sponsor',
  'SponsorEmail',
  'ProjectManager',
  'ProjectManagerEmail',
  'StartDate',
  'EndDate',
  'Status',
  'Priority',
  'RAG',
  'Progress',
  'BudgetTotal',
  'BudgetCommitted',
  'BudgetConsumed',
  'StrategicArea',
  'Notes',
];

const DELIVERABLE_COLUMNS = [
  'ProjectCode',
  'Title',
  'Owner',
  'OwnerEmail',
  'StartDate',
  'EndDate',
  'Progress',
  'Status',
  'Weight',
  'Notes',
];

const ISSUE_COLUMNS = [
  'ProjectCode',
  'Title',
  'Description',
  'Severity',
  'Owner',
  'OwnerEmail',
  'OpenDate',
  'DueDate',
  'Status',
  'Action',
  'EscalationRequired',
  'Notes',
];

const RESOURCE_COLUMNS = ['ResourceCode', 'Title', 'PersonEmail', 'Role', 'Unit', 'Capacity', 'Active', 'Notes'];

const ALLOCATION_COLUMNS = ['ProjectCode', 'ResourceCode', 'AllocationPercent', 'RoleOnProject', 'StartDate', 'EndDate', 'Notes'];

const COST_COLUMNS = ['ProjectCode', 'CostCategory', 'AmountAllocated', 'AmountCommitted', 'AmountConsumed', 'Supplier', 'Notes'];

interface ISheetDefinition {
  name: SheetName;
  columns: string[];
  dateColumns: string[];
  exampleRow: (string | number)[];
}

const SHEET_DEFINITIONS: ISheetDefinition[] = [
  {
    name: 'Projects',
    columns: PROJECT_COLUMNS,
    dateColumns: ['StartDate', 'EndDate'],
    exampleRow: [
      '#ESEMPIO',
      'Progetto di esempio',
      'Descrizione di esempio',
      'Sponsor di esempio',
      '',
      'Mario Rossi',
      'mario.rossi@contoso.com',
      '01/01/2025',
      '31/12/2025',
      'In Progress',
      'Medium',
      'Green',
      50,
      100000,
      60000,
      40000,
      'Area di esempio',
      'Note di esempio',
    ],
  },
  {
    name: 'Deliverables',
    columns: DELIVERABLE_COLUMNS,
    dateColumns: ['StartDate', 'EndDate'],
    exampleRow: [
      '#ESEMPIO',
      'Deliverable di esempio',
      'Owner di esempio',
      '',
      '01/01/2025',
      '31/03/2025',
      50,
      'In Progress',
      20,
      'Note di esempio',
    ],
  },
  {
    name: 'Issues',
    columns: ISSUE_COLUMNS,
    dateColumns: ['OpenDate', 'DueDate'],
    exampleRow: [
      '#ESEMPIO',
      'Issue di esempio',
      'Descrizione di esempio',
      'Medium',
      'Owner di esempio',
      '',
      '01/01/2025',
      '31/01/2025',
      'Open',
      'Azione di esempio',
      'No',
      'Note di esempio',
    ],
  },
  {
    name: 'Resources',
    columns: RESOURCE_COLUMNS,
    dateColumns: [],
    exampleRow: ['#ESEMPIO', 'Risorsa di esempio', '', 'Developer', 'IT', 100, 'Si', 'Note di esempio'],
  },
  {
    name: 'Allocations',
    columns: ALLOCATION_COLUMNS,
    dateColumns: ['StartDate', 'EndDate'],
    exampleRow: ['#ESEMPIO', 'RES-001', 50, 'Developer', '01/01/2025', '31/12/2025', 'Note di esempio'],
  },
  {
    name: 'Costs',
    columns: COST_COLUMNS,
    dateColumns: [],
    exampleRow: ['#ESEMPIO', 'Licenses', 10000, 8000, 5000, 'Fornitore di esempio', 'Note di esempio'],
  },
];

const SHEET_NAMES: SheetName[] = ['Projects', 'Deliverables', 'Issues', 'Resources', 'Allocations', 'Costs'];

// ---------------------------------------------------------------------------
// Chiavi di upsert
// ---------------------------------------------------------------------------

function projectKey(row: Pick<IProject, 'ProjectCode'>): string {
  return row.ProjectCode.trim();
}
function resourceKey(row: Pick<IResource, 'ResourceCode'>): string {
  return row.ResourceCode.trim();
}
function deliverableKey(row: Pick<IDeliverable, 'ProjectCode' | 'Title'>): string {
  return `${row.ProjectCode.trim()}::${row.Title.trim()}`;
}
function issueKey(row: Pick<IIssue, 'ProjectCode' | 'Title'>): string {
  return `${row.ProjectCode.trim()}::${row.Title.trim()}`;
}
function allocationKey(row: Pick<IAllocation, 'ProjectCode' | 'ResourceCode'>): string {
  return `${row.ProjectCode.trim()}::${row.ResourceCode.trim()}`;
}
function costKey(row: Pick<ICost, 'ProjectCode' | 'CostCategory'>): string {
  return `${row.ProjectCode.trim()}::${row.CostCategory.trim()}`;
}

// ---------------------------------------------------------------------------
// Parsing/formattazione di basso livello (solo API ES5: niente .find(),
// .includes(), Array.from(), Number.isNaN() - vedi tsconfig.json "lib").
// ---------------------------------------------------------------------------

function readText(row: Record<string, unknown>, column: string): string {
  const raw = row[column];
  if (typeof raw === 'string') {
    return raw.trim();
  }
  if (raw === undefined || raw === null) {
    return '';
  }
  return String(raw).trim();
}

// Usa readText (non un controllo rigido su typeof) per riconoscere il marcatore:
// readText converte in stringa qualunque valore non nullo (numero, booleano,
// ecc.), coerente con come v.requiredText/v.optionalText leggono la stessa
// cella. Un controllo piu' rigido qui avrebbe potuto restituire 'false' per una
// cella che i validatori dei campi leggono comunque come testo valido, lasciando
// passare la riga di esempio come record reale invece di scartarla.
function isCommentRow(row: Record<string, unknown>, keyColumn: string): boolean {
  const text = readText(row, keyColumn);
  return text.length > 0 && text.charAt(0) === '#';
}

function toIsoDate(date: Date): string {
  return date.toISOString();
}

function parseDateValue(raw: unknown): string | undefined {
  if (raw instanceof Date) {
    return isNaN(raw.getTime()) ? undefined : toIsoDate(raw);
  }
  if (typeof raw === 'string') {
    const trimmed = raw.trim();
    const isoMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
    if (isoMatch) {
      const year = parseInt(isoMatch[1], 10);
      const month = parseInt(isoMatch[2], 10);
      const day = parseInt(isoMatch[3], 10);
      const date = new Date(Date.UTC(year, month - 1, day));
      return isNaN(date.getTime()) ? undefined : toIsoDate(date);
    }
    const euMatch = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(trimmed);
    if (euMatch) {
      const day = parseInt(euMatch[1], 10);
      const month = parseInt(euMatch[2], 10);
      const year = parseInt(euMatch[3], 10);
      const date = new Date(Date.UTC(year, month - 1, day));
      return isNaN(date.getTime()) ? undefined : toIsoDate(date);
    }
    return undefined;
  }
  return undefined;
}

function formatDateForExport(raw: unknown): string {
  const value = typeof raw === 'string' ? raw : undefined;
  if (!value) {
    return '';
  }
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return '';
  }
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
}

function parseNumericValue(raw: unknown): number | undefined {
  if (typeof raw === 'number') {
    return isNaN(raw) ? undefined : raw;
  }
  if (typeof raw === 'string') {
    const normalized = raw.trim().replace(',', '.');
    if (normalized === '') {
      return undefined;
    }
    const value = parseFloat(normalized);
    return isNaN(value) ? undefined : value;
  }
  return undefined;
}

function parseBooleanValue(raw: unknown): boolean | undefined {
  if (typeof raw === 'boolean') {
    return raw;
  }
  if (typeof raw === 'number') {
    if (raw === 1) {
      return true;
    }
    if (raw === 0) {
      return false;
    }
    return undefined;
  }
  if (typeof raw === 'string') {
    const normalized = raw.trim().toLowerCase();
    if (normalized === 'si' || normalized === 'sì' || normalized === 'yes' || normalized === 'true' || normalized === '1') {
      return true;
    }
    if (normalized === 'no' || normalized === 'false' || normalized === '0') {
      return false;
    }
    return undefined;
  }
  return undefined;
}

function pushAll<T>(target: T[], source: T[]): void {
  source.forEach((item) => target.push(item));
}

// TKey e' il tipo minimo richiesto dalla funzione-chiave (es. Pick<IProject,'ProjectCode'>);
// TItem e' il tipo reale degli elementi (IProject o NewProject, entrambi soddisfano TKey).
// Due type parameter distinti evitano che l'inferenza di un unico T, chiamato sia dalla
// posizione covariante (items) sia da quella contravariante (keyFn), collassi sul vincolo.
function buildKeySet<TKey, TItem extends TKey>(items: TItem[], keyFn: (item: TKey) => string): Set<string> {
  const set = new Set<string>();
  items.forEach((item) => set.add(keyFn(item)));
  return set;
}

function unionSets(a: Set<string>, b: Set<string>): Set<string> {
  const result = new Set<string>();
  a.forEach((value) => result.add(value));
  b.forEach((value) => result.add(value));
  return result;
}

function buildExportFileName(): string {
  const now = new Date();
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  const datePart = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}`;
  const timePart = `${pad(now.getHours())}${pad(now.getMinutes())}`;
  return `PMO-GISO-Export-${datePart}-${timePart}.xlsx`;
}

// ---------------------------------------------------------------------------
// Validatore di riga: raccoglie gli IValidationIssue mentre legge le colonne.
// ---------------------------------------------------------------------------

class RowValidator {
  private readonly issues: IValidationIssue[] = [];

  constructor(
    private readonly sheet: SheetName,
    private readonly rowNumber: number,
  ) {}

  public requiredText(row: Record<string, unknown>, column: string): string | undefined {
    const text = readText(row, column);
    if (!text) {
      this.error(column, `Colonna obbligatoria '${column}' mancante o vuota.`);
      return undefined;
    }
    return text;
  }

  public optionalText(row: Record<string, unknown>, column: string): string | undefined {
    const text = readText(row, column);
    return text ? text : undefined;
  }

  public requiredChoice(row: Record<string, unknown>, column: string, options: string[]): string | undefined {
    const value = this.requiredText(row, column);
    if (value === undefined || value === null) {
      return undefined;
    }
    if (options.indexOf(value) === -1) {
      this.error(column, `Valore '${value}' non ammesso per '${column}'. Valori validi: ${options.join(', ')}.`);
      return undefined;
    }
    return value;
  }

  public optionalDate(row: Record<string, unknown>, column: string): string | undefined {
    const raw = row[column];
    if (raw === undefined || raw === null || raw === '') {
      return undefined;
    }
    const parsed = parseDateValue(raw);
    if (parsed === undefined || parsed === null) {
      this.error(column, `Data non valida in '${column}': '${String(raw)}'. Formati ammessi: gg/mm/aaaa oppure aaaa-mm-gg.`);
      return undefined;
    }
    return parsed;
  }

  public optionalNumber(row: Record<string, unknown>, column: string): number | undefined {
    const raw = row[column];
    if (raw === undefined || raw === null || raw === '') {
      return undefined;
    }
    const parsed = parseNumericValue(raw);
    if (parsed === undefined || parsed === null) {
      this.error(column, `Numero non valido in '${column}': '${String(raw)}'.`);
      return undefined;
    }
    return parsed;
  }

  public optionalPercent(row: Record<string, unknown>, column: string): number | undefined {
    const value = this.optionalNumber(row, column);
    if (value === undefined || value === null) {
      return undefined;
    }
    if (value < 0 || value > 100) {
      this.error(column, `Il valore di '${column}' deve essere compreso tra 0 e 100 (trovato ${value}).`);
      return undefined;
    }
    return value;
  }

  public optionalBoolean(row: Record<string, unknown>, column: string): boolean | undefined {
    const raw = row[column];
    if (raw === undefined || raw === null || raw === '') {
      return undefined;
    }
    const parsed = parseBooleanValue(raw);
    if (parsed === undefined || parsed === null) {
      this.error(column, `Valore non valido in '${column}': '${String(raw)}'. Usa Si/No, Yes/No, true/false, 1/0.`);
      return undefined;
    }
    return parsed;
  }

  public checkReference(column: string, value: string, knownKeys: Set<string>, entityLabel: string): boolean {
    if (!knownKeys.has(value.trim())) {
      this.error(column, `${entityLabel} '${value}' non trovato (ne' tra i dati gia' presenti, ne' tra quelli importati in questo file).`);
      return false;
    }
    return true;
  }

  public warnEmpty(fields: Record<string, unknown>): void {
    const emptyFields: string[] = [];
    Object.keys(fields).forEach((key) => {
      if (fields[key] === undefined || fields[key] === null) {
        emptyFields.push(key);
      }
    });
    if (emptyFields.length > 0) {
      this.warning(undefined, `Campi opzionali vuoti: ${emptyFields.join(', ')}.`);
    }
  }

  public warnDuplicate(entityLabel: string, key: string): void {
    this.warning(undefined, `${entityLabel} '${key}' duplicato nel file: la riga precedente con la stessa chiave verra' sovrascritta.`);
  }

  public warnUpdate(entityLabel: string, key: string): void {
    this.warning(undefined, `${entityLabel} '${key}' esiste gia': il record verra' aggiornato invece che creato.`);
  }

  public error(column: string | undefined, message: string): void {
    this.issues.push({ sheet: this.sheet, row: this.rowNumber, column, level: 'error', message });
  }

  public warning(column: string | undefined, message: string): void {
    this.issues.push({ sheet: this.sheet, row: this.rowNumber, column, level: 'warning', message });
  }

  public getIssues(): IValidationIssue[] {
    return this.issues;
  }
}

// ---------------------------------------------------------------------------
// Servizio
// ---------------------------------------------------------------------------

export class ExcelService {
  private readonly dataService: SharePointDataService;

  private readonly peopleService: PeopleService;

  constructor(context: WebPartContext) {
    this.dataService = new SharePointDataService(context);
    this.peopleService = new PeopleService(context);
  }

  // Risolve una colonna *Email opzionale in un Id utente numerico. Se la colonna
  // e' vuota non fa nulla (nessuna chiamata di rete, nessun campo persona
  // valorizzato). Se l'email non e' risolvibile, aggiunge un warning non
  // bloccante e la riga procede valorizzando solo il campo testo corrispondente
  // (gia' letto separatamente dal chiamante).
  private async resolvePersonEmail(
    v: RowValidator,
    raw: Record<string, unknown>,
    emailColumn: string,
  ): Promise<number | undefined> {
    const email = v.optionalText(raw, emailColumn);
    if (!email) {
      return undefined;
    }

    try {
      const resolved = await this.peopleService.ensureUserByEmail(email);
      if (!resolved) {
        v.warning(emailColumn, `Utente non trovato per l'email '${email}': verra' salvato solo il campo testo corrispondente.`);
        return undefined;
      }
      return resolved.id;
    } catch (error) {
      v.warning(
        emailColumn,
        `Impossibile risolvere l'email '${email}' (${error instanceof Error ? error.message : 'errore imprevisto'}): ` +
          'verra\' salvato solo il campo testo corrispondente.',
      );
      return undefined;
    }
  }

  // -------------------------------------------------------------- Template

  public generateTemplate(): void {
    const workbook = XLSX.utils.book_new();
    SHEET_DEFINITIONS.forEach((definition) => {
      const worksheet = XLSX.utils.aoa_to_sheet([definition.columns, definition.exampleRow]);
      XLSX.utils.book_append_sheet(workbook, worksheet, definition.name);
    });
    this.downloadWorkbook(workbook, 'PMO-GISO-Template.xlsx');
  }

  // ------------------------------------------------------------------ Import

  public async parseWorkbook(file: File): Promise<IParsedWorkbook> {
    let buffer: ArrayBuffer;
    try {
      buffer = await file.arrayBuffer();
    } catch {
      throw new ExcelParseError('Impossibile leggere il file selezionato.');
    }

    let workbook: Workbook;
    try {
      workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    } catch {
      throw new ExcelParseError("Il file selezionato non e' un file Excel valido.");
    }

    const missingSheets: SheetName[] = [];
    const sheets = {} as Record<SheetName, Record<string, unknown>[]>;

    SHEET_NAMES.forEach((name) => {
      const worksheet = workbook.Sheets[name];
      if (!worksheet) {
        missingSheets.push(name);
        sheets[name] = [];
        return;
      }
      sheets[name] = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { defval: null });
    });

    return { sheets, missingSheets };
  }

  public async validateData(parsed: IParsedWorkbook): Promise<IValidationReport> {
    // Sola lettura: nessuna scrittura in questa fase.
    const [existingProjects, existingResources, existingDeliverables, existingIssues, existingAllocations, existingCosts] =
      await Promise.all([
        this.dataService.getProjects(),
        this.dataService.getResources(),
        this.dataService.getAllDeliverables(),
        this.dataService.getAllIssues(),
        this.dataService.getAllocations(),
        this.dataService.getAllCosts(),
      ]);

    const existingProjectKeys = buildKeySet(existingProjects, projectKey);
    const existingResourceKeys = buildKeySet(existingResources, resourceKey);
    const existingDeliverableKeys = buildKeySet(existingDeliverables, deliverableKey);
    const existingIssueKeys = buildKeySet(existingIssues, issueKey);
    const existingAllocationKeys = buildKeySet(existingAllocations, allocationKey);
    const existingCostKeys = buildKeySet(existingCosts, costKey);

    const projectsResult = await this.validateProjectsSheet(parsed.sheets.Projects, existingProjectKeys);
    const projectCodesInBatch = buildKeySet(
      projectsResult.validRows.map((r) => r.data),
      projectKey,
    );
    // Unione: progetti gia' in SharePoint + progetti validi nello stesso file (verranno
    // creati per primi, vedi ordine di elaborazione in importData).
    const knownProjectCodes = unionSets(existingProjectKeys, projectCodesInBatch);

    const resourcesResult = await this.validateResourcesSheet(parsed.sheets.Resources, existingResourceKeys);
    const resourceCodesInBatch = buildKeySet(
      resourcesResult.validRows.map((r) => r.data),
      resourceKey,
    );
    const knownResourceCodes = unionSets(existingResourceKeys, resourceCodesInBatch);

    const deliverablesResult = await this.validateDeliverablesSheet(
      parsed.sheets.Deliverables,
      existingDeliverableKeys,
      knownProjectCodes,
    );
    const issuesResult = await this.validateIssuesSheet(parsed.sheets.Issues, existingIssueKeys, knownProjectCodes);
    const allocationsResult = this.validateAllocationsSheet(
      parsed.sheets.Allocations,
      existingAllocationKeys,
      knownProjectCodes,
      knownResourceCodes,
    );
    const costsResult = this.validateCostsSheet(parsed.sheets.Costs, existingCostKeys, knownProjectCodes);

    const allResults = [projectsResult, resourcesResult, deliverablesResult, issuesResult, allocationsResult, costsResult];
    let hasBlockingErrors = false;
    allResults.forEach((result) => {
      if (result.errors.length > 0) {
        hasBlockingErrors = true;
      }
    });

    return {
      Projects: projectsResult,
      Deliverables: deliverablesResult,
      Issues: issuesResult,
      Resources: resourcesResult,
      Allocations: allocationsResult,
      Costs: costsResult,
      hasBlockingErrors,
      missingSheets: parsed.missingSheets,
    };
  }

  public async importData(report: IValidationReport, callbacks?: IImportCallbacks): Promise<IImportReport> {
    const total =
      report.Projects.validRows.length +
      report.Resources.validRows.length +
      report.Deliverables.validRows.length +
      report.Issues.validRows.length +
      report.Allocations.validRows.length +
      report.Costs.validRows.length;
    const progress = { completed: 0, total };

    // Ordine obbligatorio per integrita' referenziale.
    const existingProjects = await this.dataService.getProjects();
    const projectsOutcome = await this.upsertSheet(
      'Projects',
      report.Projects.validRows,
      projectKey,
      buildIdMap(existingProjects, projectKey),
      (row) => this.dataService.createProject(row),
      (id, row) => this.dataService.updateProject(id, row),
      progress,
      callbacks,
    );

    const existingResources = await this.dataService.getResources();
    const resourcesOutcome = await this.upsertSheet(
      'Resources',
      report.Resources.validRows,
      resourceKey,
      buildIdMap(existingResources, resourceKey),
      (row) => this.dataService.createResource(row),
      (id, row) => this.dataService.updateResource(id, row),
      progress,
      callbacks,
    );

    const existingDeliverables = await this.dataService.getAllDeliverables();
    const deliverablesOutcome = await this.upsertSheet(
      'Deliverables',
      report.Deliverables.validRows,
      deliverableKey,
      buildIdMap(existingDeliverables, deliverableKey),
      (row) => this.dataService.createDeliverable(row),
      (id, row) => this.dataService.updateDeliverable(id, row),
      progress,
      callbacks,
    );

    const existingIssues = await this.dataService.getAllIssues();
    const issuesOutcome = await this.upsertSheet(
      'Issues',
      report.Issues.validRows,
      issueKey,
      buildIdMap(existingIssues, issueKey),
      (row) => this.dataService.createIssue(row),
      (id, row) => this.dataService.updateIssue(id, row),
      progress,
      callbacks,
    );

    const existingAllocations = await this.dataService.getAllocations();
    const allocationsOutcome = await this.upsertSheet(
      'Allocations',
      report.Allocations.validRows,
      allocationKey,
      buildIdMap(existingAllocations, allocationKey),
      (row) => this.dataService.createAllocation(row),
      (id, row) => this.dataService.updateAllocation(id, row),
      progress,
      callbacks,
    );

    const existingCosts = await this.dataService.getAllCosts();
    const costsOutcome = await this.upsertSheet(
      'Costs',
      report.Costs.validRows,
      costKey,
      buildIdMap(existingCosts, costKey),
      (row) => this.dataService.createCost(row),
      (id, row) => this.dataService.updateCost(id, row),
      progress,
      callbacks,
    );

    return {
      Projects: projectsOutcome,
      Resources: resourcesOutcome,
      Deliverables: deliverablesOutcome,
      Issues: issuesOutcome,
      Allocations: allocationsOutcome,
      Costs: costsOutcome,
    };
  }

  // ------------------------------------------------------------------ Export

  public async exportAllData(): Promise<void> {
    const [projects, deliverables, issues, resources, allocations, costs] = await Promise.all([
      this.dataService.getProjects(),
      this.dataService.getAllDeliverables(),
      this.dataService.getAllIssues(),
      this.dataService.getResources(),
      this.dataService.getAllocations(),
      this.dataService.getAllCosts(),
    ]);

    // I campi persona si leggono come oggetto annidato (SponsorUser.EMail, non
    // una proprieta' piatta 'SponsorEmail'): li si affianca esplicitamente prima
    // di mappare le righe sulle colonne del foglio, cosi' l'export include le
    // email delle persone collegate.
    const projectsForExport = projects.map((project) => ({
      ...project,
      SponsorEmail: project.SponsorUser?.EMail ?? '',
      ProjectManagerEmail: project.ProjectManagerUser?.EMail ?? '',
    }));
    const deliverablesForExport = deliverables.map((deliverable) => ({
      ...deliverable,
      OwnerEmail: deliverable.OwnerUser?.EMail ?? '',
    }));
    const issuesForExport = issues.map((issue) => ({ ...issue, OwnerEmail: issue.OwnerUser?.EMail ?? '' }));
    const resourcesForExport = resources.map((resource) => ({
      ...resource,
      PersonEmail: resource.PersonUser?.EMail ?? '',
    }));

    const dataBySheet: Record<SheetName, Record<string, unknown>[]> = {
      Projects: projectsForExport as unknown as Record<string, unknown>[],
      Deliverables: deliverablesForExport as unknown as Record<string, unknown>[],
      Issues: issuesForExport as unknown as Record<string, unknown>[],
      Resources: resourcesForExport as unknown as Record<string, unknown>[],
      Allocations: allocations as unknown as Record<string, unknown>[],
      Costs: costs as unknown as Record<string, unknown>[],
    };

    const workbook = XLSX.utils.book_new();
    SHEET_DEFINITIONS.forEach((definition) => {
      const rows = dataBySheet[definition.name];
      const aoa: unknown[][] = [definition.columns];
      rows.forEach((row) => {
        const line = definition.columns.map((column) => {
          const value = row[column];
          if (definition.dateColumns.indexOf(column) !== -1) {
            return formatDateForExport(value);
          }
          if (typeof value === 'boolean') {
            return value ? 'Si' : 'No';
          }
          return value === undefined || value === null ? '' : value;
        });
        aoa.push(line);
      });
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(workbook, worksheet, definition.name);
    });

    this.downloadWorkbook(workbook, buildExportFileName());
  }

  // -------------------------------------------------------- Validatori foglio

  private async validateProjectsSheet(
    rows: Record<string, unknown>[],
    existingKeys: Set<string>,
  ): Promise<ISheetValidationResult<NewProject>> {
    const issues: IValidationIssue[] = [];
    const validRows: IValidatedRow<NewProject>[] = [];
    const seenKeys = new Set<string>();
    const statusOptions = getChoiceOptions('PMO_Projects', 'Status');
    const priorityOptions = getChoiceOptions('PMO_Projects', 'Priority');
    const ragOptions = getChoiceOptions('PMO_Projects', 'RAG');

    for (let index = 0; index < rows.length; index += 1) {
      const raw = rows[index];
      const rowNumber = index + 2;
      if (isCommentRow(raw, 'ProjectCode')) {
        continue;
      }

      const v = new RowValidator('Projects', rowNumber);
      const projectCode = v.requiredText(raw, 'ProjectCode');
      const title = v.requiredText(raw, 'Title');
      const status = v.requiredChoice(raw, 'Status', statusOptions);
      const priority = v.requiredChoice(raw, 'Priority', priorityOptions);
      const rag = v.requiredChoice(raw, 'RAG', ragOptions);
      const description = v.optionalText(raw, 'Description');
      const sponsor = v.optionalText(raw, 'Sponsor');
      const sponsorUserId = await this.resolvePersonEmail(v, raw, 'SponsorEmail');
      const projectManager = v.optionalText(raw, 'ProjectManager');
      const projectManagerUserId = await this.resolvePersonEmail(v, raw, 'ProjectManagerEmail');
      const startDate = v.optionalDate(raw, 'StartDate');
      const endDate = v.optionalDate(raw, 'EndDate');
      const progress = v.optionalPercent(raw, 'Progress');
      const budgetTotal = v.optionalNumber(raw, 'BudgetTotal');
      const budgetCommitted = v.optionalNumber(raw, 'BudgetCommitted');
      const budgetConsumed = v.optionalNumber(raw, 'BudgetConsumed');
      const strategicArea = v.optionalText(raw, 'StrategicArea');
      const notes = v.optionalText(raw, 'Notes');

      v.warnEmpty({
        Description: description,
        Sponsor: sponsor,
        ProjectManager: projectManager,
        StartDate: startDate,
        EndDate: endDate,
        Progress: progress,
        BudgetTotal: budgetTotal,
        BudgetCommitted: budgetCommitted,
        BudgetConsumed: budgetConsumed,
        StrategicArea: strategicArea,
        Notes: notes,
      });

      if (
        projectCode !== undefined &&
        projectCode !== null &&
        title !== undefined &&
        title !== null &&
        status !== undefined &&
        status !== null &&
        priority !== undefined &&
        priority !== null &&
        rag !== undefined &&
        rag !== null
      ) {
        const key = projectCode.trim();
        if (seenKeys.has(key)) {
          v.warnDuplicate('ProjectCode', projectCode);
        }
        seenKeys.add(key);
        const isUpdate = existingKeys.has(key);
        if (isUpdate) {
          v.warnUpdate('Progetto', projectCode);
        }
        validRows.push({
          rowNumber,
          isUpdate,
          data: {
            ProjectCode: projectCode,
            Title: title,
            Status: status as ProjectStatus,
            Priority: priority as ProjectPriority,
            RAG: rag as ProjectRag,
            Description: description,
            Sponsor: sponsor,
            SponsorUserId: sponsorUserId,
            ProjectManager: projectManager,
            ProjectManagerUserId: projectManagerUserId,
            StartDate: startDate,
            EndDate: endDate,
            Progress: progress,
            BudgetTotal: budgetTotal,
            BudgetCommitted: budgetCommitted,
            BudgetConsumed: budgetConsumed,
            StrategicArea: strategicArea,
            Notes: notes,
          },
        });
      }

      pushAll(issues, v.getIssues());
    }

    return {
      totalRows: rows.length,
      validRows,
      errors: issues.filter((i) => i.level === 'error'),
      warnings: issues.filter((i) => i.level === 'warning'),
    };
  }

  private async validateResourcesSheet(
    rows: Record<string, unknown>[],
    existingKeys: Set<string>,
  ): Promise<ISheetValidationResult<NewResource>> {
    const issues: IValidationIssue[] = [];
    const validRows: IValidatedRow<NewResource>[] = [];
    const seenKeys = new Set<string>();

    for (let index = 0; index < rows.length; index += 1) {
      const raw = rows[index];
      const rowNumber = index + 2;
      if (isCommentRow(raw, 'ResourceCode')) {
        continue;
      }

      const v = new RowValidator('Resources', rowNumber);
      const resourceCode = v.requiredText(raw, 'ResourceCode');
      const title = v.requiredText(raw, 'Title');
      const personUserId = await this.resolvePersonEmail(v, raw, 'PersonEmail');
      const role = v.optionalText(raw, 'Role');
      const unit = v.optionalText(raw, 'Unit');
      const capacity = v.optionalNumber(raw, 'Capacity');
      const active = v.optionalBoolean(raw, 'Active');
      const notes = v.optionalText(raw, 'Notes');

      v.warnEmpty({ Role: role, Unit: unit, Capacity: capacity, Active: active, Notes: notes });

      if (resourceCode !== undefined && resourceCode !== null && title !== undefined && title !== null) {
        const key = resourceCode.trim();
        if (seenKeys.has(key)) {
          v.warnDuplicate('ResourceCode', resourceCode);
        }
        seenKeys.add(key);
        const isUpdate = existingKeys.has(key);
        if (isUpdate) {
          v.warnUpdate('Risorsa', resourceCode);
        }
        validRows.push({
          rowNumber,
          isUpdate,
          data: {
            ResourceCode: resourceCode,
            Title: title,
            PersonUserId: personUserId,
            Role: role,
            Unit: unit,
            Capacity: capacity,
            Active: active,
            Notes: notes,
          },
        });
      }

      pushAll(issues, v.getIssues());
    }

    return {
      totalRows: rows.length,
      validRows,
      errors: issues.filter((i) => i.level === 'error'),
      warnings: issues.filter((i) => i.level === 'warning'),
    };
  }

  private async validateDeliverablesSheet(
    rows: Record<string, unknown>[],
    existingKeys: Set<string>,
    knownProjectCodes: Set<string>,
  ): Promise<ISheetValidationResult<NewDeliverable>> {
    const issues: IValidationIssue[] = [];
    const validRows: IValidatedRow<NewDeliverable>[] = [];
    const seenKeys = new Set<string>();
    const statusOptions = getChoiceOptions('PMO_Deliverables', 'Status');

    for (let index = 0; index < rows.length; index += 1) {
      const raw = rows[index];
      const rowNumber = index + 2;
      if (isCommentRow(raw, 'ProjectCode')) {
        continue;
      }

      const v = new RowValidator('Deliverables', rowNumber);
      const projectCode = v.requiredText(raw, 'ProjectCode');
      const title = v.requiredText(raw, 'Title');
      const status = v.requiredChoice(raw, 'Status', statusOptions);
      const owner = v.optionalText(raw, 'Owner');
      const ownerUserId = await this.resolvePersonEmail(v, raw, 'OwnerEmail');
      const startDate = v.optionalDate(raw, 'StartDate');
      const endDate = v.optionalDate(raw, 'EndDate');
      const progress = v.optionalPercent(raw, 'Progress');
      const weight = v.optionalNumber(raw, 'Weight');
      const notes = v.optionalText(raw, 'Notes');

      v.warnEmpty({ Owner: owner, StartDate: startDate, EndDate: endDate, Progress: progress, Weight: weight, Notes: notes });

      let projectExists = false;
      if (projectCode !== undefined && projectCode !== null) {
        projectExists = v.checkReference('ProjectCode', projectCode, knownProjectCodes, 'Progetto');
      }

      if (
        projectCode !== undefined &&
        projectCode !== null &&
        title !== undefined &&
        title !== null &&
        status !== undefined &&
        status !== null &&
        projectExists
      ) {
        const key = `${projectCode.trim()}::${title.trim()}`;
        if (seenKeys.has(key)) {
          v.warnDuplicate('Deliverable', `${projectCode} / ${title}`);
        }
        seenKeys.add(key);
        const isUpdate = existingKeys.has(key);
        if (isUpdate) {
          v.warnUpdate('Deliverable', `${projectCode} / ${title}`);
        }
        validRows.push({
          rowNumber,
          isUpdate,
          data: {
            ProjectCode: projectCode,
            Title: title,
            Status: status as DeliverableStatus,
            Owner: owner,
            OwnerUserId: ownerUserId,
            StartDate: startDate,
            EndDate: endDate,
            Progress: progress,
            Weight: weight,
            Notes: notes,
          },
        });
      }

      pushAll(issues, v.getIssues());
    }

    return {
      totalRows: rows.length,
      validRows,
      errors: issues.filter((i) => i.level === 'error'),
      warnings: issues.filter((i) => i.level === 'warning'),
    };
  }

  private async validateIssuesSheet(
    rows: Record<string, unknown>[],
    existingKeys: Set<string>,
    knownProjectCodes: Set<string>,
  ): Promise<ISheetValidationResult<NewIssue>> {
    const issues: IValidationIssue[] = [];
    const validRows: IValidatedRow<NewIssue>[] = [];
    const seenKeys = new Set<string>();
    const severityOptions = getChoiceOptions('PMO_Issues', 'Severity');
    const statusOptions = getChoiceOptions('PMO_Issues', 'Status');

    for (let index = 0; index < rows.length; index += 1) {
      const raw = rows[index];
      const rowNumber = index + 2;
      if (isCommentRow(raw, 'ProjectCode')) {
        continue;
      }

      const v = new RowValidator('Issues', rowNumber);
      const projectCode = v.requiredText(raw, 'ProjectCode');
      const title = v.requiredText(raw, 'Title');
      const severity = v.requiredChoice(raw, 'Severity', severityOptions);
      const status = v.requiredChoice(raw, 'Status', statusOptions);
      const description = v.optionalText(raw, 'Description');
      const owner = v.optionalText(raw, 'Owner');
      const ownerUserId = await this.resolvePersonEmail(v, raw, 'OwnerEmail');
      const openDate = v.optionalDate(raw, 'OpenDate');
      const dueDate = v.optionalDate(raw, 'DueDate');
      const action = v.optionalText(raw, 'Action');
      const escalationRequired = v.optionalBoolean(raw, 'EscalationRequired');
      const notes = v.optionalText(raw, 'Notes');

      v.warnEmpty({
        Description: description,
        Owner: owner,
        OpenDate: openDate,
        DueDate: dueDate,
        Action: action,
        EscalationRequired: escalationRequired,
        Notes: notes,
      });

      let projectExists = false;
      if (projectCode !== undefined && projectCode !== null) {
        projectExists = v.checkReference('ProjectCode', projectCode, knownProjectCodes, 'Progetto');
      }

      if (
        projectCode !== undefined &&
        projectCode !== null &&
        title !== undefined &&
        title !== null &&
        severity !== undefined &&
        severity !== null &&
        status !== undefined &&
        status !== null &&
        projectExists
      ) {
        const key = `${projectCode.trim()}::${title.trim()}`;
        if (seenKeys.has(key)) {
          v.warnDuplicate('Issue', `${projectCode} / ${title}`);
        }
        seenKeys.add(key);
        const isUpdate = existingKeys.has(key);
        if (isUpdate) {
          v.warnUpdate('Issue', `${projectCode} / ${title}`);
        }
        validRows.push({
          rowNumber,
          isUpdate,
          data: {
            ProjectCode: projectCode,
            Title: title,
            Severity: severity as IssueSeverity,
            Status: status as IssueStatus,
            Description: description,
            Owner: owner,
            OwnerUserId: ownerUserId,
            OpenDate: openDate,
            DueDate: dueDate,
            Action: action,
            EscalationRequired: escalationRequired,
            Notes: notes,
          },
        });
      }

      pushAll(issues, v.getIssues());
    }

    return {
      totalRows: rows.length,
      validRows,
      errors: issues.filter((i) => i.level === 'error'),
      warnings: issues.filter((i) => i.level === 'warning'),
    };
  }

  private validateAllocationsSheet(
    rows: Record<string, unknown>[],
    existingKeys: Set<string>,
    knownProjectCodes: Set<string>,
    knownResourceCodes: Set<string>,
  ): ISheetValidationResult<NewAllocation> {
    const issues: IValidationIssue[] = [];
    const validRows: IValidatedRow<NewAllocation>[] = [];
    const seenKeys = new Set<string>();

    rows.forEach((raw, index) => {
      const rowNumber = index + 2;
      if (isCommentRow(raw, 'ProjectCode')) {
        return;
      }

      const v = new RowValidator('Allocations', rowNumber);
      const projectCode = v.requiredText(raw, 'ProjectCode');
      const resourceCode = v.requiredText(raw, 'ResourceCode');
      const allocationPercent = v.optionalPercent(raw, 'AllocationPercent');
      const roleOnProject = v.optionalText(raw, 'RoleOnProject');
      const startDate = v.optionalDate(raw, 'StartDate');
      const endDate = v.optionalDate(raw, 'EndDate');
      const notes = v.optionalText(raw, 'Notes');

      v.warnEmpty({ AllocationPercent: allocationPercent, RoleOnProject: roleOnProject, StartDate: startDate, EndDate: endDate, Notes: notes });

      let projectExists = false;
      if (projectCode !== undefined && projectCode !== null) {
        projectExists = v.checkReference('ProjectCode', projectCode, knownProjectCodes, 'Progetto');
      }
      let resourceExists = false;
      if (resourceCode !== undefined && resourceCode !== null) {
        resourceExists = v.checkReference('ResourceCode', resourceCode, knownResourceCodes, 'Risorsa');
      }

      if (
        projectCode !== undefined &&
        projectCode !== null &&
        resourceCode !== undefined &&
        resourceCode !== null &&
        projectExists &&
        resourceExists
      ) {
        const key = `${projectCode.trim()}::${resourceCode.trim()}`;
        if (seenKeys.has(key)) {
          v.warnDuplicate('Allocazione', `${projectCode} / ${resourceCode}`);
        }
        seenKeys.add(key);
        const isUpdate = existingKeys.has(key);
        if (isUpdate) {
          v.warnUpdate('Allocazione', `${projectCode} / ${resourceCode}`);
        }
        validRows.push({
          rowNumber,
          isUpdate,
          data: {
            ProjectCode: projectCode,
            ResourceCode: resourceCode,
            Title: `${projectCode}-${resourceCode}`,
            AllocationPercent: allocationPercent,
            RoleOnProject: roleOnProject,
            StartDate: startDate,
            EndDate: endDate,
            Notes: notes,
          },
        });
      }

      pushAll(issues, v.getIssues());
    });

    return {
      totalRows: rows.length,
      validRows,
      errors: issues.filter((i) => i.level === 'error'),
      warnings: issues.filter((i) => i.level === 'warning'),
    };
  }

  private validateCostsSheet(
    rows: Record<string, unknown>[],
    existingKeys: Set<string>,
    knownProjectCodes: Set<string>,
  ): ISheetValidationResult<NewCost> {
    const issues: IValidationIssue[] = [];
    const validRows: IValidatedRow<NewCost>[] = [];
    const seenKeys = new Set<string>();
    const categoryOptions = getChoiceOptions('PMO_Costs', 'CostCategory');

    rows.forEach((raw, index) => {
      const rowNumber = index + 2;
      if (isCommentRow(raw, 'ProjectCode')) {
        return;
      }

      const v = new RowValidator('Costs', rowNumber);
      const projectCode = v.requiredText(raw, 'ProjectCode');
      const costCategory = v.requiredChoice(raw, 'CostCategory', categoryOptions);
      const amountAllocated = v.optionalNumber(raw, 'AmountAllocated');
      const amountCommitted = v.optionalNumber(raw, 'AmountCommitted');
      const amountConsumed = v.optionalNumber(raw, 'AmountConsumed');
      const supplier = v.optionalText(raw, 'Supplier');
      const notes = v.optionalText(raw, 'Notes');

      v.warnEmpty({ AmountAllocated: amountAllocated, AmountCommitted: amountCommitted, AmountConsumed: amountConsumed, Supplier: supplier, Notes: notes });

      let projectExists = false;
      if (projectCode !== undefined && projectCode !== null) {
        projectExists = v.checkReference('ProjectCode', projectCode, knownProjectCodes, 'Progetto');
      }

      if (projectCode !== undefined && projectCode !== null && costCategory !== undefined && costCategory !== null && projectExists) {
        const key = `${projectCode.trim()}::${costCategory.trim()}`;
        if (seenKeys.has(key)) {
          v.warnDuplicate('Costo', `${projectCode} / ${costCategory}`);
        }
        seenKeys.add(key);
        const isUpdate = existingKeys.has(key);
        if (isUpdate) {
          v.warnUpdate('Costo', `${projectCode} / ${costCategory}`);
        }
        validRows.push({
          rowNumber,
          isUpdate,
          data: {
            ProjectCode: projectCode,
            CostCategory: costCategory as CostCategory,
            Title: `${projectCode}-${costCategory}`,
            AmountAllocated: amountAllocated,
            AmountCommitted: amountCommitted,
            AmountConsumed: amountConsumed,
            Supplier: supplier,
            Notes: notes,
          },
        });
      }

      pushAll(issues, v.getIssues());
    });

    return {
      totalRows: rows.length,
      validRows,
      errors: issues.filter((i) => i.level === 'error'),
      warnings: issues.filter((i) => i.level === 'warning'),
    };
  }

  // -------------------------------------------------------------- Internals

  // TKey e' il tipo minimo richiesto dalla funzione-chiave (es. Pick<IProject,'ProjectCode'>);
  // TNew e' il tipo reale della riga da creare/aggiornare (NewProject, NewResource, ecc.),
  // vincolato a soddisfare TKey. Due type parameter distinti evitano che l'inferenza di un
  // unico T collassi sul vincolo quando viene usato sia in posizione covariante (rows,
  // create, update) sia in posizione contravariante (getKey) - vedi buildKeySet/buildIdMap.
  private async upsertSheet<TKey, TNew extends TKey>(
    sheet: SheetName,
    rows: IValidatedRow<TNew>[],
    getKey: (row: TKey) => string,
    existingKeyToId: Map<string, number>,
    create: (row: TNew) => Promise<IWriteResult>,
    update: (id: number, row: TNew) => Promise<IWriteResult>,
    progress: { completed: number; total: number },
    callbacks?: IImportCallbacks,
  ): Promise<ISheetImportResult> {
    let created = 0;
    let updated = 0;
    let skipped = 0;
    const errors: IValidationIssue[] = [];

    for (const row of rows) {
      const key = getKey(row.data);
      const existingId = existingKeyToId.get(key);

      try {
        const result =
          existingId !== undefined && existingId !== null
            ? await update(existingId, row.data)
            : await create(row.data);
        if (result.success) {
          if (existingId !== undefined && existingId !== null) {
            updated += 1;
            callbacks?.onLog?.({ sheet, row: row.rowNumber, level: 'success', message: `Aggiornato: ${result.message}` });
          } else {
            created += 1;
            if (result.id !== undefined && result.id !== null) {
              existingKeyToId.set(key, result.id);
            }
            callbacks?.onLog?.({ sheet, row: row.rowNumber, level: 'success', message: `Creato: ${result.message}` });
          }
        } else {
          skipped += 1;
          errors.push({ sheet, row: row.rowNumber, level: 'error', message: result.message });
          callbacks?.onLog?.({ sheet, row: row.rowNumber, level: 'error', message: result.message });
        }
      } catch (error) {
        skipped += 1;
        const message = error instanceof Error ? error.message : 'Errore imprevisto.';
        errors.push({ sheet, row: row.rowNumber, level: 'error', message });
        callbacks?.onLog?.({ sheet, row: row.rowNumber, level: 'error', message });
      }

      progress.completed += 1;
      callbacks?.onProgress?.(progress.completed, progress.total);
    }

    return { created, updated, skipped, errors };
  }

  private downloadWorkbook(workbook: Workbook, fileName: string): void {
    const wbout: unknown = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout as BlobPart], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }
}

// Stessa correzione di buildKeySet: TKey per la funzione-chiave, TItem (con Id) per gli
// elementi reali letti da SharePoint.
function buildIdMap<TKey, TItem extends TKey & { Id: number }>(
  items: TItem[],
  keyFn: (item: TKey) => string,
): Map<string, number> {
  const map = new Map<string, number>();
  items.forEach((item) => map.set(keyFn(item), item.Id));
  return map;
}
