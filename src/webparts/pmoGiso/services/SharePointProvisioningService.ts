import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

export type FieldType = 'Text' | 'Note' | 'DateTime' | 'Number' | 'Currency' | 'Choice' | 'Boolean' | 'User';

export interface IFieldSchema {
  internalName: string;
  type: FieldType;
  choices?: string[];
}

export interface IListSchema {
  title: string;
  description: string;
  fields: IFieldSchema[];
}

export type EnvironmentItemStatus = 'Existing' | 'Created' | 'Missing' | 'Error';

export interface IFieldStatus {
  internalName: string;
  status: EnvironmentItemStatus;
  message: string;
}

export interface IListStatus {
  listTitle: string;
  status: EnvironmentItemStatus;
  message: string;
  fields: IFieldStatus[];
}

export interface IEnvironmentReport {
  results: IListStatus[];
  createdCount: number;
  existingCount: number;
  errorCount: number;
}

const PERMISSION_DENIED_MESSAGE =
  "Permessi insufficienti per creare liste SharePoint. Serve il ruolo Owner del sito o permessi Manage Lists.";

// BaseTemplate 100 = Custom List.
const CUSTOM_LIST_TEMPLATE = 100;

// Titolo dell'elemento temporaneo usato da verifyUserFieldUsable: riconoscibile
// nel caso (eccezionale) in cui la cancellazione fallisca e la riga resti visibile.
const VERIFICATION_ITEM_TITLE = '__PMO_GISO_VERIFICA_CAMPO_ELIMINABILE__';

export const REQUIRED_LISTS: IListSchema[] = [
  {
    title: 'PMO_Projects',
    description: 'Anagrafica progetti PMO GISO.',
    fields: [
      { internalName: 'ProjectCode', type: 'Text' },
      { internalName: 'Description', type: 'Note' },
      { internalName: 'Sponsor', type: 'Text' },
      // Colonne Persona: affiancano i campi testo esistenti (mantenuti per
      // soggetti esterni al tenant e per compatibilita' con i dati gia' inseriti),
      // non li sostituiscono.
      { internalName: 'SponsorUser', type: 'User' },
      // Storicizza l'email dell'utente selezionato in SponsorUser al momento della
      // selezione: a differenza del campo User, un testo libero non dipende dalla
      // presenza dell'account nel tenant e resta leggibile anche se l'identita'
      // cambia email o lascia il tenant.
      { internalName: 'SponsorUserEmail', type: 'Text' },
      { internalName: 'ProjectManager', type: 'Text' },
      { internalName: 'ProjectManagerUser', type: 'User' },
      { internalName: 'ProjectManagerUserEmail', type: 'Text' },
      { internalName: 'StartDate', type: 'DateTime' },
      { internalName: 'EndDate', type: 'DateTime' },
      {
        internalName: 'Status',
        type: 'Choice',
        choices: ['Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled'],
      },
      { internalName: 'Priority', type: 'Choice', choices: ['Low', 'Medium', 'High', 'Critical'] },
      { internalName: 'RAG', type: 'Choice', choices: ['Green', 'Amber', 'Red', 'Grey'] },
      { internalName: 'Progress', type: 'Number' },
      { internalName: 'BudgetTotal', type: 'Currency' },
      { internalName: 'BudgetCommitted', type: 'Currency' },
      { internalName: 'BudgetConsumed', type: 'Currency' },
      { internalName: 'StrategicArea', type: 'Text' },
      { internalName: 'Notes', type: 'Note' },
    ],
  },
  {
    title: 'PMO_Deliverables',
    description: 'Deliverable/milestone dei progetti PMO GISO.',
    fields: [
      { internalName: 'ProjectCode', type: 'Text' },
      { internalName: 'Owner', type: 'Text' },
      { internalName: 'OwnerUser', type: 'User' },
      { internalName: 'OwnerUserEmail', type: 'Text' },
      { internalName: 'StartDate', type: 'DateTime' },
      { internalName: 'EndDate', type: 'DateTime' },
      { internalName: 'Progress', type: 'Number' },
      {
        internalName: 'Status',
        type: 'Choice',
        choices: ['Not Started', 'In Progress', 'Completed', 'Delayed', 'Cancelled'],
      },
      { internalName: 'Weight', type: 'Number' },
      { internalName: 'Notes', type: 'Note' },
    ],
  },
  {
    title: 'PMO_Issues',
    description: 'Issue/rischi dei progetti PMO GISO.',
    fields: [
      { internalName: 'ProjectCode', type: 'Text' },
      { internalName: 'Description', type: 'Note' },
      { internalName: 'Severity', type: 'Choice', choices: ['Low', 'Medium', 'High', 'Critical'] },
      { internalName: 'Owner', type: 'Text' },
      { internalName: 'OwnerUser', type: 'User' },
      { internalName: 'OwnerUserEmail', type: 'Text' },
      { internalName: 'OpenDate', type: 'DateTime' },
      { internalName: 'DueDate', type: 'DateTime' },
      { internalName: 'Status', type: 'Choice', choices: ['Open', 'In Progress', 'Mitigated', 'Closed'] },
      { internalName: 'Action', type: 'Note' },
      { internalName: 'EscalationRequired', type: 'Boolean' },
      { internalName: 'Notes', type: 'Note' },
    ],
  },
  {
    title: 'PMO_Resources',
    description: 'Anagrafica risorse PMO GISO.',
    fields: [
      { internalName: 'ResourceCode', type: 'Text' },
      { internalName: 'Role', type: 'Text' },
      { internalName: 'Unit', type: 'Text' },
      { internalName: 'Capacity', type: 'Number' },
      { internalName: 'Active', type: 'Boolean' },
      { internalName: 'PersonUser', type: 'User' },
      { internalName: 'PersonUserEmail', type: 'Text' },
      { internalName: 'Notes', type: 'Note' },
    ],
  },
  {
    title: 'PMO_Allocations',
    description: 'Allocazione delle risorse sui progetti PMO GISO.',
    fields: [
      { internalName: 'ProjectCode', type: 'Text' },
      { internalName: 'ResourceCode', type: 'Text' },
      { internalName: 'AllocationPercent', type: 'Number' },
      { internalName: 'RoleOnProject', type: 'Text' },
      { internalName: 'StartDate', type: 'DateTime' },
      { internalName: 'EndDate', type: 'DateTime' },
      { internalName: 'Notes', type: 'Note' },
    ],
  },
  {
    title: 'PMO_Costs',
    description: 'Costi dei progetti PMO GISO.',
    fields: [
      { internalName: 'ProjectCode', type: 'Text' },
      {
        internalName: 'CostCategory',
        type: 'Choice',
        choices: [
          'External Consulting',
          'Internal Consulting',
          'HW/SW Purchase',
          'Internal Effort',
          'Licenses',
          'Other',
        ],
      },
      { internalName: 'AmountAllocated', type: 'Currency' },
      { internalName: 'AmountCommitted', type: 'Currency' },
      { internalName: 'AmountConsumed', type: 'Currency' },
      { internalName: 'Supplier', type: 'Text' },
      { internalName: 'Notes', type: 'Note' },
    ],
  },
];

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Gli apici singoli nei literal OData (es. GetByTitle('...')) vanno raddoppiati.
function escapeODataLiteral(value: string): string {
  return value.replace(/'/g, "''");
}

function buildFieldSchemaXml(field: IFieldSchema): string {
  const name = escapeXml(field.internalName);
  const commonAttrs = `Type="${field.type}" Name="${name}" StaticName="${name}" DisplayName="${name}"`;

  if (field.type === 'Choice') {
    const choices = field.choices ?? [];
    const choicesXml = choices.map((choice) => `<CHOICE>${escapeXml(choice)}</CHOICE>`).join('');
    const defaultXml = choices.length > 0 ? `<Default>${escapeXml(choices[0])}</Default>` : '';
    return `<Field ${commonAttrs}>${defaultXml}<CHOICES>${choicesXml}</CHOICES></Field>`;
  }

  if (field.type === 'Note') {
    return `<Field ${commonAttrs} RichText="FALSE" />`;
  }

  if (field.type === 'User') {
    // Tutti gli attributi richiesti da un campo Persona valido, non solo Type e
    // List: un XML incompleto puo' produrre un campo che sembra corretto nei
    // metadati (InternalName/TypeAsString/LookupList tutti a posto) ma che di
    // fatto non e' utilizzabile (query e scritture falliscono). Mult="FALSE" =
    // valore singolo, coerente con l'uso previsto (Sponsor, PM, Owner, risorsa).
    return (
      `<Field ${commonAttrs} List="UserInfo" ShowField="ImageName" ` +
      `UserSelectionMode="PeopleOnly" UserSelectionScope="0" Mult="FALSE" Required="FALSE" />`
    );
  }

  return `<Field ${commonAttrs} />`;
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

export class SharePointProvisioningService {
  private readonly spHttpClient: SPHttpClient;

  private readonly webUrl: string;

  constructor(context: WebPartContext) {
    this.spHttpClient = context.spHttpClient;
    this.webUrl = context.pageContext.web.absoluteUrl;
  }

  public async checkListExists(listTitle: string): Promise<boolean> {
    const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')?$select=Title`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (response.status === 404) {
      return false;
    }
    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }
    return true;
  }

  public async createList(
    listTitle: string,
    description: string,
  ): Promise<{ success: boolean; message: string }> {
    const url = `${this.webUrl}/_api/web/lists`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        // SPHttpClient aggiunge di default 'odata-version: 4.0', incompatibile con
        // payload __metadata in formato odata=verbose (v3): senza questo azzeramento
        // SharePoint non riesce a determinare l'entity set e rifiuta la richiesta.
        'odata-version': '',
      },
      body: JSON.stringify({
        __metadata: { type: 'SP.List' },
        Title: listTitle,
        Description: description,
        BaseTemplate: CUSTOM_LIST_TEMPLATE,
      }),
    });

    if (!response.ok) {
      return { success: false, message: await this.describeErrorResponse(response) };
    }
    return { success: true, message: `Lista '${listTitle}' creata.` };
  }

  public async ensureField(listTitle: string, field: IFieldSchema): Promise<IFieldStatus> {
    const existing = await this.checkFieldExists(listTitle, field.internalName);
    if (!existing.success) {
      return { internalName: field.internalName, status: 'Error', message: existing.message };
    }
    if (existing.exists) {
      return { internalName: field.internalName, status: 'Existing', message: 'Colonna gia\' presente.' };
    }

    const created = await this.createField(listTitle, field);
    if (!created.success) {
      return { internalName: field.internalName, status: 'Error', message: created.message };
    }

    // Un campo Persona puo' risultare creato con metadati apparentemente
    // corretti (InternalName/TypeAsString/LookupList a posto) ma di fatto
    // inutilizzabile: si rilegge e si esercita davvero prima di dichiararlo
    // creato con successo, altrimenti si segnala come Error.
    if (field.type === 'User') {
      const verified = await this.verifyUserFieldUsable(listTitle, field.internalName);
      if (!verified.success) {
        return {
          internalName: field.internalName,
          status: 'Error',
          message: `Colonna '${field.internalName}' creata ma non utilizzabile: ${verified.message}`,
        };
      }
    }

    return { internalName: field.internalName, status: 'Created', message: `Colonna '${field.internalName}' creata.` };
  }

  public async checkEnvironment(): Promise<IListStatus[]> {
    const results: IListStatus[] = [];

    for (const list of REQUIRED_LISTS) {
      results.push(await this.checkListStatus(list));
    }

    return results;
  }

  public async initializeEnvironment(): Promise<IEnvironmentReport> {
    const results: IListStatus[] = [];

    for (const list of REQUIRED_LISTS) {
      results.push(await this.initializeList(list));
    }

    let createdCount = 0;
    let existingCount = 0;
    let errorCount = 0;
    for (const list of results) {
      if (list.status === 'Created') createdCount += 1;
      else if (list.status === 'Existing') existingCount += 1;
      else if (list.status === 'Error') errorCount += 1;

      for (const field of list.fields) {
        if (field.status === 'Created') createdCount += 1;
        else if (field.status === 'Existing') existingCount += 1;
        else if (field.status === 'Error') errorCount += 1;
      }
    }

    return { results, createdCount, existingCount, errorCount };
  }

  private async checkListStatus(list: IListSchema): Promise<IListStatus> {
    let exists: boolean;
    try {
      exists = await this.checkListExists(list.title);
    } catch (error) {
      return {
        listTitle: list.title,
        status: 'Error',
        message: error instanceof Error ? error.message : 'Errore durante la verifica della lista.',
        fields: list.fields.map((field) => ({
          internalName: field.internalName,
          status: 'Error',
          message: 'Verifica non eseguita: la lista non e\' raggiungibile.',
        })),
      };
    }

    if (!exists) {
      return {
        listTitle: list.title,
        status: 'Missing',
        message: 'Lista non trovata.',
        fields: list.fields.map((field) => ({
          internalName: field.internalName,
          status: 'Missing',
          message: 'La lista non esiste ancora.',
        })),
      };
    }

    const fields: IFieldStatus[] = [];
    for (const field of list.fields) {
      const fieldExists = await this.checkFieldExists(list.title, field.internalName);
      fields.push(
        fieldExists.success
          ? {
              internalName: field.internalName,
              status: fieldExists.exists ? 'Existing' : 'Missing',
              message: fieldExists.exists ? 'Colonna presente.' : 'Colonna non trovata.',
            }
          : { internalName: field.internalName, status: 'Error', message: fieldExists.message },
      );
    }

    return { listTitle: list.title, status: 'Existing', message: 'Lista presente.', fields };
  }

  private async initializeList(list: IListSchema): Promise<IListStatus> {
    let exists: boolean;
    try {
      exists = await this.checkListExists(list.title);
    } catch (error) {
      return {
        listTitle: list.title,
        status: 'Error',
        message: error instanceof Error ? error.message : 'Errore durante la verifica della lista.',
        fields: list.fields.map((field) => ({
          internalName: field.internalName,
          status: 'Error',
          message: 'Colonna non elaborata: la lista non e\' raggiungibile.',
        })),
      };
    }

    let listStatus: EnvironmentItemStatus;
    let listMessage: string;

    if (exists) {
      listStatus = 'Existing';
      listMessage = 'Lista gia\' presente.';
    } else {
      const created = await this.createList(list.title, list.description);
      if (!created.success) {
        return {
          listTitle: list.title,
          status: 'Error',
          message: created.message,
          fields: list.fields.map((field) => ({
            internalName: field.internalName,
            status: 'Error',
            message: 'Colonna non elaborata: creazione della lista fallita.',
          })),
        };
      }
      listStatus = 'Created';
      listMessage = created.message;
    }

    const fields: IFieldStatus[] = [];
    for (const field of list.fields) {
      fields.push(await this.ensureField(list.title, field));
    }

    return { listTitle: list.title, status: listStatus, message: listMessage, fields };
  }

  private async checkFieldExists(
    listTitle: string,
    internalName: string,
  ): Promise<{ success: boolean; exists: boolean; message: string }> {
    const url =
      `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/fields` +
      `?$filter=InternalName eq '${escapeODataLiteral(internalName)}'&$select=Id`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (response.status === 404) {
      return { success: false, exists: false, message: 'Lista non trovata durante la verifica della colonna.' };
    }
    if (!response.ok) {
      return { success: false, exists: false, message: await this.describeErrorResponse(response) };
    }

    const body = (await this.safeJson(response)) as { value?: unknown[] };
    const values = Array.isArray(body.value) ? body.value : [];
    return { success: true, exists: values.length > 0, message: '' };
  }

  private async createField(
    listTitle: string,
    field: IFieldSchema,
  ): Promise<{ success: boolean; message: string }> {
    const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/fields/createfieldasxml`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        // Vedi commento in createList: azzera l'header 'odata-version: 4.0'
        // aggiunto di default da SPHttpClient, incompatibile con payload odata=verbose.
        'odata-version': '',
      },
      body: JSON.stringify({
        parameters: {
          __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
          SchemaXml: buildFieldSchemaXml(field),
        },
      }),
    });

    if (!response.ok) {
      return { success: false, message: await this.describeErrorResponse(response) };
    }
    return { success: true, message: `Colonna '${field.internalName}' creata.` };
  }

  // Verifica che un campo Persona appena creato sia davvero utilizzabile.
  // Non usa $select=Id,<Nome>Id su /items: su una lista appena creata (quindi
  // senza elementi) quella query puo' restituire HTTP 500 "Input string was
  // not in a correct format" per un difetto noto del motore REST di SharePoint,
  // indipendente dalla reale validita' del campo. Si esercita invece un vero
  // ciclo scrittura/cancellazione (lo stesso path di un salvataggio reale,
  // dove il problema originale si manifestava anche nell'interfaccia nativa):
  // si crea un elemento temporaneo valorizzando il campo con l'utente
  // corrente, si verifica che la scrittura riesca, poi lo si elimina sempre,
  // cosi' il provisioning resta idempotente e non lascia righe fittizie.
  private async verifyUserFieldUsable(listTitle: string, internalName: string): Promise<{ success: boolean; message: string }> {
    let entityType: string;
    let currentUserId: number;
    try {
      entityType = await this.getListItemEntityTypeFullName(listTitle);
      currentUserId = await this.getCurrentUserId();
    } catch (error) {
      return { success: false, message: error instanceof Error ? error.message : 'Verifica non eseguita.' };
    }

    const createUrl = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/items`;
    const createResponse = await this.spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify({
        __metadata: { type: entityType },
        Title: VERIFICATION_ITEM_TITLE,
        [`${internalName}Id`]: currentUserId,
      }),
    });

    if (!createResponse.ok) {
      return { success: false, message: await this.describeErrorResponse(createResponse) };
    }

    const createdBody = (await this.safeJson(createResponse)) as { d?: { Id?: number } };
    const itemId = createdBody.d?.Id;
    if (itemId != null) {
      await this.deleteVerificationItem(listTitle, itemId);
    }

    return { success: true, message: '' };
  }

  private async getListItemEntityTypeFullName(listTitle: string): Promise<string> {
    const url =
      `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')` +
      `?$select=ListItemEntityTypeFullName`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as { ListItemEntityTypeFullName?: string };
    const entityType = body.ListItemEntityTypeFullName;
    if (!entityType) {
      throw new Error(`Impossibile determinare il tipo di elemento per la lista '${listTitle}'.`);
    }
    return entityType;
  }

  private async getCurrentUserId(): Promise<number> {
    const url = `${this.webUrl}/_api/web/currentuser?$select=Id`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as { Id?: number };
    if (body.Id == null) {
      throw new Error("Impossibile determinare l'utente corrente.");
    }
    return body.Id;
  }

  // Cancellazione best-effort: un fallimento qui non deve far dichiarare
  // 'Error' un campo che si e' invece dimostrato scrivibile: nel caso peggiore
  // resta una riga di verifica riconoscibile (VERIFICATION_ITEM_TITLE) da
  // rimuovere manualmente, preferibile a bloccare l'intero provisioning.
  private async deleteVerificationItem(listTitle: string, itemId: number): Promise<void> {
    try {
      const url = `${this.webUrl}/_api/web/lists/GetByTitle('${escapeODataLiteral(listTitle)}')/items(${itemId})`;
      await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': '',
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*',
        },
      });
    } catch {
      // Ignorato volutamente, vedi commento sopra.
    }
  }

  private async describeErrorResponse(response: SPHttpClientResponse): Promise<string> {
    if (response.status === 403) {
      return PERMISSION_DENIED_MESSAGE;
    }

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
