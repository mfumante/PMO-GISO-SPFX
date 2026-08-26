import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPersonSearchResult {
  displayName: string;
  email: string;
  loginName: string;
}

export interface IResolvedUser {
  id: number;
  displayName: string;
  email: string;
  loginName: string;
}

// Voce della mappa Id -> utente costruita da getSiteUsers(), usata da
// SharePointDataService per arricchire lato client i campi persona letti come
// semplice *Id (niente $expand, vedi getSiteUsers).
export interface ISiteUserInfo {
  id: number;
  displayName: string;
  email: string;
}

interface IPickerEntityData {
  Email?: string;
}

interface IPickerEntity {
  Key: string;
  DisplayText: string;
  IsResolved: boolean;
  EntityData?: IPickerEntityData;
}

// Header comuni a tutte le scritture (POST con __metadata in formato odata=verbose).
// SPHttpClient aggiunge di default 'odata-version: 4.0', incompatibile con questo
// formato: va azzerato, vedi SharePointProvisioningService.ts.
const WRITE_HEADERS: Record<string, string> = {
  Accept: 'application/json;odata=verbose',
  'Content-Type': 'application/json;odata=verbose',
  'odata-version': '',
};

// PrincipalType 1 = solo utenti reali (esclude gruppi di SharePoint, distribution
// list, security group): coerente con l'uso previsto (Sponsor, PM, Owner, risorsa).
const PRINCIPAL_TYPE_USER = 1;
// PrincipalSource 15 = tutte le origini (UserInfoList, Search, MembershipProvider,
// RoleProvider): il valore di default raccomandato da Microsoft.
const PRINCIPAL_SOURCE_ALL = 15;
const MAX_SUGGESTIONS = 20;

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

export class PeopleService {
  private readonly spHttpClient: SPHttpClient;

  private readonly webUrl: string;

  // Cache in memoria delle risoluzioni gia' effettuate in questa sessione, per
  // evitare chiamate ripetute a ensureuser per la stessa persona. Chiavi separate
  // per login name (usate da ensureUser) ed email (usate da ensureUserByEmail),
  // per evitare collisioni fra i due spazi di chiavi nella stessa mappa.
  private readonly resolveCache = new Map<string, IResolvedUser>();

  // Cache degli utenti del sito (vedi getSiteUsers): popolata alla prima
  // chiamata riuscita e riusata per tutta la vita di questa istanza, cosi'
  // SharePointDataService non deve rileggerla per ogni entita'/riga.
  private siteUsersCache: Map<number, ISiteUserInfo> | undefined;

  constructor(context: WebPartContext) {
    this.spHttpClient = context.spHttpClient;
    this.webUrl = context.pageContext.web.absoluteUrl;
  }

  // Legge una sola volta l'elenco utenti del sito (senza $expand: e' una lista
  // piatta, sempre disponibile) e lo mette in cache come mappa Id -> info,
  // usata per arricchire lato client i campi persona letti come *Id.
  public async getSiteUsers(): Promise<Map<number, ISiteUserInfo>> {
    if (this.siteUsersCache) {
      return this.siteUsersCache;
    }

    const url = `${this.webUrl}/_api/web/siteusers?$select=Id,Title,Email,LoginName&$top=5000`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as { value?: { Id: number; Title?: string; Email?: string }[] };
    const rawUsers = Array.isArray(body.value) ? body.value : [];

    const map = new Map<number, ISiteUserInfo>();
    rawUsers.forEach((user) => {
      map.set(user.Id, { id: user.Id, displayName: user.Title ?? '', email: user.Email ?? '' });
    });

    this.siteUsersCache = map;
    return map;
  }

  public async searchUsers(query: string): Promise<IPersonSearchResult[]> {
    const trimmed = query.trim();
    if (!trimmed) {
      return [];
    }

    const url = `${this.webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: WRITE_HEADERS,
      body: JSON.stringify({
        queryParams: {
          __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: MAX_SUGGESTIONS,
          PrincipalSource: PRINCIPAL_SOURCE_ALL,
          PrincipalType: PRINCIPAL_TYPE_USER,
          QueryString: trimmed,
        },
      }),
    });

    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as { d?: { ClientPeoplePickerSearchUser?: string } };
    const raw = body.d?.ClientPeoplePickerSearchUser;
    if (!raw) {
      return [];
    }

    let entities: IPickerEntity[];
    try {
      const parsed = JSON.parse(raw) as unknown;
      entities = Array.isArray(parsed) ? (parsed as IPickerEntity[]) : [];
    } catch {
      return [];
    }

    return entities
      .filter((entity) => entity.IsResolved)
      .map((entity) => ({
        displayName: entity.DisplayText,
        email: entity.EntityData?.Email ?? '',
        loginName: entity.Key,
      }));
  }

  public async ensureUser(loginName: string): Promise<IResolvedUser> {
    const cached = this.resolveCache.get(loginName);
    if (cached) {
      return cached;
    }

    const url = `${this.webUrl}/_api/web/ensureuser`;
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: WRITE_HEADERS,
      body: JSON.stringify({ logonName: loginName }),
    });

    if (!response.ok) {
      throw new Error(await this.describeErrorResponse(response));
    }

    const body = (await this.safeJson(response)) as {
      d?: { Id?: number; Title?: string; Email?: string; LoginName?: string };
    };
    const id = body.d?.Id;
    if (id == null) {
      throw new Error(`Impossibile risolvere l'utente '${loginName}'.`);
    }

    const resolved: IResolvedUser = {
      id,
      displayName: body.d?.Title ?? loginName,
      email: body.d?.Email ?? '',
      loginName: body.d?.LoginName ?? loginName,
    };
    this.resolveCache.set(loginName, resolved);
    return resolved;
  }

  // Risoluzione per indirizzo email, usata dall'import Excel (le colonne
  // *Email dei fogli non contengono un login name pronto per ensureuser):
  // cerca l'utente via people picker e lo risolve poi con ensureUser.
  public async ensureUserByEmail(email: string): Promise<IResolvedUser | undefined> {
    const trimmed = email.trim();
    if (!trimmed) {
      return undefined;
    }

    const cacheKey = `email:${trimmed.toLowerCase()}`;
    const cached = this.resolveCache.get(cacheKey);
    if (cached) {
      return cached;
    }

    const results = await this.searchUsers(trimmed);
    const match = results.filter((result) => result.email.toLowerCase() === trimmed.toLowerCase())[0];
    if (!match) {
      return undefined;
    }

    const resolved = await this.ensureUser(match.loginName);
    this.resolveCache.set(cacheKey, resolved);
    return resolved;
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
