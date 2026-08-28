import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { REQUIRED_LISTS } from './SharePointProvisioningService';
import { PeopleService } from './PeopleService';
import type { ISiteUserInfo } from './PeopleService';

// ---------------------------------------------------------------------------
// Tipi Choice, coerenti con gli schemi definiti in SharePointProvisioningService.
// ---------------------------------------------------------------------------

export type ProjectStatus = 'Not Started' | 'In Progress' | 'Completed' | 'On Hold' | 'Cancelled';
export type ProjectPriority = 'Low' | 'Medium' | 'High' | 'Critical';
export type ProjectRag = 'Green' | 'Amber' | 'Red' | 'Grey';
export type DeliverableStatus = 'Not Started' | 'In Progress' | 'Completed' | 'Delayed' | 'Cancelled';
export type IssueSeverity = 'Low' | 'Medium' | 'High' | 'Critical';
export type IssueStatus = 'Open' | 'In Progress' | 'Mitigated' | 'Closed';
export type CostCategory =
  | 'External Consulting'
  | 'Internal Consulting'
  | 'HW/SW Purchase'
  | 'Internal Effort'
  | 'Licenses'
  | 'Other';

// Riferimento a un utente SharePoint per un campo Persona (People). Costruito
// lato client a partire dal campo piatto '<Nome>Id' (mai da $expand, vedi
// attachPersonFields) incrociato con la mappa utenti di PeopleService.getSiteUsers().
// Popolato solo quando il campo e' valorizzato e la colonna User esiste davvero
// nella lista.
export interface IPersonFieldRef {
  Id?: number;
  Title?: string;
  EMail?: string;
}

export interface IProject {
  Id: number;
  Title: string;
  ProjectCode: string;
  Description?: string;
  Sponsor?: string;
  SponsorUser?: IPersonFieldRef;
  // Email dell'utente selezionato in SponsorUser, storicizzata al momento della
  // selezione (vedi PeopleService/EntityFormDialog syncEmailField): a differenza
  // di SponsorUser.EMail (risolta live da PeopleService.getSiteUsers()), questo
  // valore resta invariato anche se l'identita' cambia email o lascia il tenant.
  SponsorUserEmail?: string;
  ProjectManager?: string;
  ProjectManagerUser?: IPersonFieldRef;
  ProjectManagerUserEmail?: string;
  StartDate?: string;
  EndDate?: string;
  Status: ProjectStatus;
  Priority: ProjectPriority;
  RAG: ProjectRag;
  Progress?: number;
  BudgetTotal?: number;
  BudgetCommitted?: number;
  BudgetConsumed?: number;
  StrategicArea?: string;
  Notes?: string;
}

export interface IDeliverable {
  Id: number;
  Title: string;
  ProjectCode: string;
  Owner?: string;
  OwnerUser?: IPersonFieldRef;
  // Vedi commento su IProject.SponsorUserEmail.
  OwnerUserEmail?: string;
  StartDate?: string;
  EndDate?: string;
  Progress?: number;
  Status: DeliverableStatus;
  Weight?: number;
  Notes?: string;
}

export interface IIssue {
  Id: number;
  Title: string;
  ProjectCode: string;
  Description?: string;
  Severity: IssueSeverity;
  Owner?: string;
  OwnerUser?: IPersonFieldRef;
  // Vedi commento su IProject.SponsorUserEmail.
  OwnerUserEmail?: string;
  OpenDate?: string;
  DueDate?: string;
  Status: IssueStatus;
  Action?: string;
  EscalationRequired?: boolean;
  Notes?: string;
}

export interface IResource {
  Id: number;
  Title: string;
  ResourceCode: string;
  Role?: string;
  Unit?: string;
  Capacity?: number;
  Active?: boolean;
  PersonUser?: IPersonFieldRef;
  // Vedi commento su IProject.SponsorUserEmail.
  PersonUserEmail?: string;
  Notes?: string;
}

export interface IAllocation {
  Id: number;
  Title: string;
  ProjectCode: string;
  ResourceCode: string;
  AllocationPercent?: number;
  RoleOnProject?: string;
  StartDate?: string;
  EndDate?: string;
  Notes?: string;
}

export interface ICost {
  Id: number;
  Title: string;
  ProjectCode: string;
  CostCategory: CostCategory;
  AmountAllocated?: number;
  AmountCommitted?: number;
  AmountConsumed?: number;
  Supplier?: string;
  Notes?: string;
}

// I campi Persona si leggono come oggetto annidato (IPersonFieldRef, costruito
// lato client da attachPersonFields) ma si scrivono valorizzando la proprieta'
// con suffisso Id (numero) o null per svuotare il campo: le due forme vanno
// tenute distinte.
export type NewProject = Omit<IProject, 'Id' | 'SponsorUser' | 'ProjectManagerUser'> & {
  // null (non undefined) e' richiesto dalla REST API di SharePoint per svuotare esplicitamente il campo persona.
  // eslint-disable-next-line @rushstack/no-new-null
  SponsorUserId?: number | null;
  // eslint-disable-next-line @rushstack/no-new-null
  ProjectManagerUserId?: number | null;
};
// eslint-disable-next-line @rushstack/no-new-null
export type NewDeliverable = Omit<IDeliverable, 'Id' | 'OwnerUser'> & { OwnerUserId?: number | null };
// eslint-disable-next-line @rushstack/no-new-null
export type NewIssue = Omit<IIssue, 'Id' | 'OwnerUser'> & { OwnerUserId?: number | null };
// eslint-disable-next-line @rushstack/no-new-null
export type NewResource = Omit<IResource, 'Id' | 'PersonUser'> & { PersonUserId?: number | null };
export type NewAllocation = Omit<IAllocation, 'Id'>;
export type NewCost = Omit<ICost, 'Id'>;

export interface IWriteResult {
  success: boolean;
  message: string;
  // Popolato solo da createItem: serve a chi importa piu' righe nella stessa
  // sessione (es. ExcelService) per riconoscere un elemento appena creato senza
  // rileggere la lista, cosi' una chiave duplicata nello stesso file aggiorna
  // invece di duplicare.
  id?: number;
}

// Lanciato quando una lista richiesta non esiste ancora (HTTP 404): permette alla
// UI di distinguere "lista da inizializzare" da un vero errore, senza far crashare
// il componente.
export class SharePointListNotFoundError extends Error {
  constructor(public readonly listTitle: string) {
    super(`La lista '${listTitle}' non esiste.`);
    this.name = 'SharePointListNotFoundError';
  }
}

const LIST_TITLES = {
  Projects: 'PMO_Projects',
  Deliverables: 'PMO_Deliverables',
  Issues: 'PMO_Issues',
  Resources: 'PMO_Resources',
  Allocations: 'PMO_Allocations',
  Costs: 'PMO_Costs',
} as const;

// Un campo Persona e la colonna Text che ne storicizza l'email (vedi
// IProject.SponsorUserEmail): entrambe le colonne vengono verificate
// dinamicamente ed entrambe possono mancare indipendentemente l'una dall'altra
// (es. subito dopo questo aggiornamento, prima di un nuovo 'Inizializza
// Ambiente'), quindi si controllano separatamente in buildSelectWithPersonFields.
interface IPersonFieldPair {
  personField: string;
  emailField: string;
}

// Niente $expand da nessuna parte (vedi buildSelectWithPersonFields/attachPersonFields):
// per ogni entita' si tengono separati i campi di base (sempre presenti) dai nomi
// delle colonne Persona, la cui esistenza viene verificata dinamicamente ad ogni
// lettura cosi' l'app resta utilizzabile anche se quelle colonne non esistono.
const PROJECT_FIELDS = [
  'Id',
  'Title',
  'ProjectCode',
  'Description',
  'Sponsor',
  'ProjectManager',
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
const PROJECT_PERSON_FIELDS: IPersonFieldPair[] = [
  { personField: 'SponsorUser', emailField: 'SponsorUserEmail' },
  { personField: 'ProjectManagerUser', emailField: 'ProjectManagerUserEmail' },
];

const DELIVERABLE_FIELDS = [
  'Id',
  'Title',
  'ProjectCode',
  'Owner',
  'StartDate',
  'EndDate',
  'Progress',
  'Status',
  'Weight',
  'Notes',
];
const DELIVERABLE_PERSON_FIELDS: IPersonFieldPair[] = [{ personField: 'OwnerUser', emailField: 'OwnerUserEmail' }];

const ISSUE_FIELDS = [
  'Id',
  'Title',
  'ProjectCode',
  'Description',
  'Severity',
  'Owner',
  'OpenDate',
  'DueDate',
  'Status',
  'Action',
  'EscalationRequired',
  'Notes',
];
const ISSUE_PERSON_FIELDS: IPersonFieldPair[] = [{ personField: 'OwnerUser', emailField: 'OwnerUserEmail' }];

const RESOURCE_FIELDS = ['Id', 'Title', 'ResourceCode', 'Role', 'Unit', 'Capacity', 'Active', 'Notes'];
const RESOURCE_PERSON_FIELDS: IPersonFieldPair[] = [{ personField: 'PersonUser', emailField: 'PersonUserEmail' }];

const ALLOCATION_FIELDS = [
  'Id',
  'Title',
  'ProjectCode',
  'ResourceCode',
  'AllocationPercent',
  'RoleOnProject',
  'StartDate',
  'EndDate',
  'Notes',
];

const COST_FIELDS = [
  'Id',
  'Title',
  'ProjectCode',
  'CostCategory',
  'AmountAllocated',
  'AmountCommitted',
  'AmountConsumed',
  'Supplier',
  'Notes',
];

// Gli apici singoli nei literal OData (es. GetByTitle('...') o eq '...') vanno raddoppiati.
function escapeODataLiteral(value: string): string {
  return value.replace(/'/g, "''");
}

function extractErrorMessage(body: unknown): string | undefined {
  const errorBody = body as { error?: { message?: string | { value?: string } } } | undefined;
  const message = errorBody?.error?.message;
  if (typeof message === 'string') {
    return message;
  }
  if (message && typeof message.value === 'string') {
    return message.value;
  }
  return undefined;
}

// Header comuni a tutte le scritture (POST con __metadata, MERGE, DELETE). SPHttpClient
// aggiunge di default 'odata-version: 4.0', incompatibile con payload __metadata in
// formato odata=verbose (v3): va azzerato, vedi SharePointProvisioningService.ts.
const WRITE_HEADERS: Record<string, string> = {
  Accept: 'application/json;odata=verbose',
  'Content-Type': 'application/json;odata=verbose',
  'odata-version': '',
};

// Restituisce le scelte configurate per un campo Choice, leggendole dagli schemi
// di provisioning invece di duplicarle a mano (restano sempre allineate).
export function getChoiceOptions(listTitle: string, fieldName: string): string[] {
  const list = REQUIRED_LISTS.filter((l) => l.title === listTitle)[0];
  const field = list?.fields.filter((f) => f.internalName === fieldName)[0];
  return field?.choices ?? [];
}

export class SharePointDataService {
  private readonly spHttpClient: SPHttpClient;

  private readonly webUrl: string;

  private readonly peopleService: PeopleService;

  private readonly entityTypeCache = new Map<string, string>();

  // Nomi delle colonne effettivamente presenti in ciascuna lista, verificati
  // dinamicamente (vedi getExistingFieldNames): solo le colonne Persona che
  // risultano davvero presenti vengono incluse nelle query.
  private readonly listFieldsCache = new Map<string, Set<string>>();

  constructor(context: WebPartContext) {
    this.spHttpClient = context.spHttpClient;
    this.webUrl = context.pageContext.web.absoluteUrl;
    this.peopleService = new PeopleService(context);
  }

  // ---------------------------------------------------------------- Projects

  public getProjects(): Promise<IProject[]> {
    return this.getItemsWithPersonFields<IProject>(LIST_TITLES.Projects, PROJECT_FIELDS, PROJECT_PERSON_FIELDS);
  }

  public async getProjectByCode(projectCode: string): Promise<IProject | undefined> {
    const items = await this.getItemsWithPersonFields<IProject>(
      LIST_TITLES.Projects,
      PROJECT_FIELDS,
      PROJECT_PERSON_FIELDS,
      `ProjectCode eq '${escapeODataLiteral(projectCode)}'`,
    );
    return items[0];
  }

  public createProject(project: NewProject): Promise<IWriteResult> {
    return this.createItem(LIST_TITLES.Projects, project);
  }

  public updateProject(itemId: number, project: Partial<NewProject>): Promise<IWriteResult> {
    return this.updateItem(LIST_TITLES.Projects, itemId, project);
  }

  public deleteProject(itemId: number): Promise<IWriteResult> {
    return this.deleteItem(LIST_TITLES.Projects, itemId);
  }

  // ----------------------------------------------------------- Deliverables

  public getAllDeliverables(): Promise<IDeliverable[]> {
    return this.getItemsWithPersonFields<IDeliverable>(
      LIST_TITLES.Deliverables,
      DELIVERABLE_FIELDS,
      DELIVERABLE_PERSON_FIELDS,
    );
  }

  public getDeliverablesByProject(projectCode: string): Promise<IDeliverable[]> {
    return this.getItemsWithPersonFields<IDeliverable>(
      LIST_TITLES.Deliverables,
      DELIVERABLE_FIELDS,
      DELIVERABLE_PERSON_FIELDS,
      `ProjectCode eq '${escapeODataLiteral(projectCode)}'`,
    );
  }

  public createDeliverable(deliverable: NewDeliverable): Promise<IWriteResult> {
    return this.createItem(LIST_TITLES.Deliverables, deliverable);
  }

  public updateDeliverable(itemId: number, deliverable: Partial<NewDeliverable>): Promise<IWriteResult> {
    return this.updateItem(LIST_TITLES.Deliverables, itemId, deliverable);
  }

  public deleteDeliverable(itemId: number): Promise<IWriteResult> {
    return this.deleteItem(LIST_TITLES.Deliverables, itemId);
  }

  // ----------------------------------------------------------------- Issues

  public getAllIssues(): Promise<IIssue[]> {
    return this.getItemsWithPersonFields<IIssue>(LIST_TITLES.Issues, ISSUE_FIELDS, ISSUE_PERSON_FIELDS);
  }

  public getIssuesByProject(projectCode: string): Promise<IIssue[]> {
    return this.getItemsWithPersonFields<IIssue>(
      LIST_TITLES.Issues,
      ISSUE_FIELDS,
      ISSUE_PERSON_FIELDS,
      `ProjectCode eq '${escapeODataLiteral(projectCode)}'`,
    );
  }

  public createIssue(issue: NewIssue): Promise<IWriteResult> {
    return this.createItem(LIST_TITLES.Issues, issue);
  }

  public updateIssue(itemId: number, issue: Partial<NewIssue>): Promise<IWriteResult> {
    return this.updateItem(LIST_TITLES.Issues, itemId, issue);
  }

  public deleteIssue(itemId: number): Promise<IWriteResult> {
    return this.deleteItem(LIST_TITLES.Issues, itemId);
  }

  // -------------------------------------------------------------- Resources

  public getResources(): Promise<IResource[]> {
    return this.getItemsWithPersonFields<IResource>(LIST_TITLES.Resources, RESOURCE_FIELDS, RESOURCE_PERSON_FIELDS);
  }

  public createResource(resource: NewResource): Promise<IWriteResult> {
    return this.createItem(LIST_TITLES.Resources, resource);
  }

  public updateResource(itemId: number, resource: Partial<NewResource>): Promise<IWriteResult> {
    return this.updateItem(LIST_TITLES.Resources, itemId, resource);
  }

  public deleteResource(itemId: number): Promise<IWriteResult> {
    return this.deleteItem(LIST_TITLES.Resources, itemId);
  }

  // ------------------------------------------------------------ Allocations

  public getAllocations(): Promise<IAllocation[]> {
    return this.getItems<IAllocation>(LIST_TITLES.Allocations, ALLOCATION_FIELDS);
  }

  public getAllocationsByProject(projectCode: string): Promise<IAllocation[]> {
    return this.getItems<IAllocation>(
      LIST_TITLES.Allocations,
      ALLOCATION_FIELDS,
      `ProjectCode eq '${escapeODataLiteral(projectCode)}'`,
    );
  }

  public getAllocationsByResource(resourceCode: string): Promise<IAllocation[]> {
    return this.getItems<IAllocation>(
      LIST_TITLES.Allocations,
      ALLOCATION_FIELDS,
      `ResourceCode eq '${escapeODataLiteral(resourceCode)}'`,
    );
  }

  public createAllocation(allocation: NewAllocation): Promise<IWriteResult> {
    return this.createItem(LIST_TITLES.Allocations, allocation);
  }

  public updateAllocation(itemId: number, allocation: Partial<NewAllocation>): Promise<IWriteResult> {
    return this.updateItem(LIST_TITLES.Allocations, itemId, allocation);
  }

  public deleteAllocation(itemId: number): Promise<IWriteResult> {
    return this.deleteItem(LIST_TITLES.Allocations, itemId);
  }

  // ----------------------------------------------------------------- Costs

  public getAllCosts(): Promise<ICost[]> {
    return this.getItems<ICost>(LIST_TITLES.Costs, COST_FIELDS);
  }

  public getCostsByProject(projectCode: string): Promise<ICost[]> {
    return this.getItems<ICost>(
      LIST_TITLES.Costs,
      COST_FIELDS,
      `ProjectCode eq '${escapeODataLiteral(projectCode)}'`,
    );
  }

  public createCost(cost: NewCost): Promise<IWriteResult> {
    return this.createItem(LIST_TITLES.Costs, cost);
  }

  public updateCost(itemId: number, cost: Partial<NewCost>): Promise<IWriteResult> {
    return this.updateItem(LIST_TITLES.Costs, itemId, cost);
  }

  public deleteCost(itemId: number): Promise<IWriteResult> {
    return this.deleteItem(LIST_TITLES.Costs, itemId);
  }

  // ------------------------------------------------------------- Internals

  private async getListItemEntityTypeFullName(listTitle: string): Promise<string> {
    const cached = this.entityTypeCache.get(listTitle);
    if (cached) {
      return cached;
    }

    const url =
      `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')` +
      `?$select=ListItemEntityTypeFullName`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (response.status === 404) {
      throw new SharePointListNotFoundError(listTitle);
    }
    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as { ListItemEntityTypeFullName?: string };
    const entityType = body.ListItemEntityTypeFullName;
    if (!entityType) {
      throw new Error(`Impossibile determinare il tipo di elemento per la lista '${listTitle}'.`);
    }

    this.entityTypeCache.set(listTitle, entityType);
    return entityType;
  }

  // Nomi delle colonne effettivamente presenti in una lista (rilevamento dinamico,
  // punto A della correzione): usato per includere nelle query solo le colonne
  // Persona che esistono davvero, cosi' l'app resta utilizzabile anche se sono
  // state rimosse o non sono ancora state (ri)create dal provisioning. Se la
  // verifica stessa fallisce per qualunque motivo, si assume prudenzialmente che
  // nessuna colonna Persona sia disponibile (nessun risultato viene messo in
  // cache in quel caso, cosi' un fallimento transitorio non resta "bloccato" per
  // tutta la vita del servizio).
  private async getExistingFieldNames(listTitle: string): Promise<Set<string>> {
    const cached = this.listFieldsCache.get(listTitle);
    if (cached) {
      return cached;
    }

    try {
      const url =
        `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/fields` +
        `?$select=InternalName&$top=5000`;
      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        return new Set<string>();
      }

      const body = (await this.safeJson(response)) as { value?: { InternalName?: string }[] };
      const rawFields = Array.isArray(body.value) ? body.value : [];

      const names = new Set<string>();
      rawFields.forEach((field) => {
        if (field.InternalName) {
          names.add(field.InternalName);
        }
      });

      this.listFieldsCache.set(listTitle, names);
      return names;
    } catch {
      return new Set<string>();
    }
  }

  // Calcola il $select da usare per un'entita': i campi di base sempre inclusi,
  // piu' '<Nome>Id' per le colonne Persona presenti (mai $expand, vedi getItems)
  // e la colonna Text che ne storicizza l'email, se presente. Le due colonne di
  // una stessa coppia si verificano indipendentemente: dopo questo aggiornamento
  // una lista puo' avere gia' il campo Persona ma non ancora la colonna Email
  // (richiede un nuovo 'Inizializza Ambiente'), e viceversa non deve mai accadere.
  private async buildSelectWithPersonFields(
    listTitle: string,
    baseFields: string[],
    personFieldPairs: IPersonFieldPair[],
  ): Promise<{ selectFields: string[]; presentPersonFields: string[] }> {
    if (personFieldPairs.length === 0) {
      return { selectFields: baseFields, presentPersonFields: [] };
    }

    const existingFields = await this.getExistingFieldNames(listTitle);
    const presentPersonFields = personFieldPairs
      .filter((pair) => existingFields.has(pair.personField))
      .map((pair) => pair.personField);
    const presentEmailFields = personFieldPairs
      .filter((pair) => existingFields.has(pair.emailField))
      .map((pair) => pair.emailField);
    const selectFields = baseFields
      .concat(presentPersonFields.map((name) => `${name}Id`))
      .concat(presentEmailFields);
    return { selectFields, presentPersonFields };
  }

  // Arricchisce lato client gli elementi gia' letti (campo piatto '<Nome>Id')
  // con l'oggetto annidato IPersonFieldRef atteso dai componenti di
  // visualizzazione esistenti (PersonDisplay e affini), senza mai chiamare
  // $expand. Se la lettura degli utenti del sito fallisce, procede comunque
  // restituendo gli elementi con i soli campi testo, senza errore.
  private async attachPersonFields<T>(items: T[], presentPersonFields: string[]): Promise<T[]> {
    if (presentPersonFields.length === 0 || items.length === 0) {
      return items;
    }

    let siteUsers: Map<number, ISiteUserInfo>;
    try {
      siteUsers = await this.peopleService.getSiteUsers();
    } catch {
      return items;
    }

    items.forEach((item) => {
      const record = item as unknown as Record<string, unknown>;
      presentPersonFields.forEach((fieldName) => {
        const idValue = record[`${fieldName}Id`];
        if (typeof idValue === 'number') {
          const user = siteUsers.get(idValue);
          if (user) {
            record[fieldName] = { Id: user.id, Title: user.displayName, EMail: user.email };
          }
        }
      });
    });

    return items;
  }

  // Legge un'entita' includendo (quando presenti) le colonne Persona, con
  // fallback resiliente: SharePoint puo' restituire HTTP 500 "Input string
  // was not in a correct format" quando il $select include la proprieta'
  // '<Nome>Id' di un campo Persona e la lista non ha (ancora) elementi
  // corrispondenti - un difetto noto del motore REST, non un'indicazione che
  // il campo sia rotto (la scrittura viene verificata a parte in
  // SharePointProvisioningService.verifyUserFieldUsable). Se la lettura con
  // le colonne Persona fallisce, si riprova senza: l'app resta utilizzabile
  // mostrando solo i campi testo invece di rompersi.
  private async getItemsWithPersonFields<T>(
    listTitle: string,
    baseFields: string[],
    personFieldPairs: IPersonFieldPair[],
    filter?: string,
  ): Promise<T[]> {
    const { selectFields, presentPersonFields } = await this.buildSelectWithPersonFields(
      listTitle,
      baseFields,
      personFieldPairs,
    );

    try {
      const items = await this.getItems<T>(listTitle, selectFields, filter);
      return await this.attachPersonFields(items, presentPersonFields);
    } catch (error) {
      if (presentPersonFields.length === 0) {
        throw error;
      }
      console.warn(
        `Lettura di '${listTitle}' con le colonne Persona (${presentPersonFields.join(', ')}) fallita, ` +
          `si riprova senza: ${error instanceof Error ? error.message : String(error)}`,
      );
      return this.getItems<T>(listTitle, baseFields, filter);
    }
  }

  // Niente $expand: i campi Persona si leggono sempre come proprieta' piatta
  // '<Nome>Id' (vedi buildSelectWithPersonFields), risolta lato client in
  // attachPersonFields. $expand su un campo Persona provisionato in modo
  // incompleto puo' rispondere 500 anche quando i metadati del campo sembrano
  // corretti; evitarlo del tutto rende l'app resiliente a colonne User assenti
  // o malformate.
  private async getItems<T>(listTitle: string, selectFields: string[], filter?: string): Promise<T[]> {
    const params = [`$select=${selectFields.join(',')}`, '$top=5000'];
    if (filter) {
      params.push(`$filter=${encodeURIComponent(filter)}`);
    }

    const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/items?${params.join('&')}`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (response.status === 404) {
      throw new SharePointListNotFoundError(listTitle);
    }
    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as { value?: T[] };
    return Array.isArray(body.value) ? body.value : [];
  }

  private async createItem<TFields extends object>(listTitle: string, fields: TFields): Promise<IWriteResult> {
    let entityType: string;
    try {
      entityType = await this.getListItemEntityTypeFullName(listTitle);
    } catch (error) {
      return { success: false, message: error instanceof Error ? error.message : 'Errore imprevisto.' };
    }

    const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/items`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: WRITE_HEADERS,
      body: JSON.stringify({ __metadata: { type: entityType }, ...fields }),
    });

    if (!response.ok) {
      return { success: false, message: await this.describeErrorResponse(response) };
    }

    // Risposta odata=verbose: l'entita' creata e' racchiusa in "d".
    const body = (await this.safeJson(response)) as { d?: { Id?: number } };
    const id = body.d?.Id;
    return { success: true, message: 'Elemento creato.', id };
  }

  private async updateItem<TFields extends object>(
    listTitle: string,
    itemId: number,
    fields: TFields,
  ): Promise<IWriteResult> {
    let entityType: string;
    try {
      entityType = await this.getListItemEntityTypeFullName(listTitle);
    } catch (error) {
      return { success: false, message: error instanceof Error ? error.message : 'Errore imprevisto.' };
    }

    const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/items(${itemId})`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        ...WRITE_HEADERS,
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
      },
      body: JSON.stringify({ __metadata: { type: entityType }, ...fields }),
    });

    if (!response.ok) {
      return { success: false, message: await this.describeErrorResponse(response) };
    }
    return { success: true, message: 'Elemento aggiornato.' };
  }

  private async deleteItem(listTitle: string, itemId: number): Promise<IWriteResult> {
    const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/items(${itemId})`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        ...WRITE_HEADERS,
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
    });

    if (!response.ok) {
      return { success: false, message: await this.describeErrorResponse(response) };
    }
    return { success: true, message: 'Elemento eliminato.' };
  }

  private async describeErrorResponse(response: SPHttpClientResponse): Promise<string> {
    const body = await this.safeJson(response);
    const message = extractErrorMessage(body);
    return `HTTP ${response.status}: ${message ?? 'richiesta fallita.'}`;
  }

  private async safeJson(response: SPHttpClientResponse): Promise<unknown> {
    try {
      const text = await response.text();
      return text ? JSON.parse(text) : {};
    } catch {
      return {};
    }
  }
}
