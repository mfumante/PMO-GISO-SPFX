import type { FormValues } from '../common/EntityFormDialog';
import type { IPersonValue } from '../common/PeoplePicker';
import type { PeopleService } from '../../services/PeopleService';
import type {
  CostCategory,
  DeliverableStatus,
  IAllocation,
  ICost,
  IDeliverable,
  IIssue,
  IPersonFieldRef,
  IProject,
  IResource,
  IssueSeverity,
  IssueStatus,
  NewAllocation,
  NewCost,
  NewDeliverable,
  NewIssue,
  NewProject,
  NewResource,
  ProjectPriority,
  ProjectRag,
  ProjectStatus,
} from '../../services/SharePointDataService';

// Converte il riferimento letto via $expand (IPersonFieldRef) nel valore atteso
// dal PeoplePicker: l'Id e' gia' noto, quindi non servira' ri-risolvere l'utente
// al salvataggio se la selezione non viene toccata (vedi resolvePersonFieldId).
function personRefToValue(ref: IPersonFieldRef | undefined): IPersonValue | undefined {
  if (!ref || ref.Id === undefined || ref.Id === null) {
    return undefined;
  }
  return { id: ref.Id, displayName: ref.Title ?? '', email: ref.EMail ?? '' };
}

// Risolve l'Id numerico da scrivere per un campo persona: se il valore ha gia'
// un Id (record esistente, non ri-selezionato) lo riusa senza chiamate di rete;
// se ha solo il loginName (nuova selezione da ricerca) lo risolve con ensureUser;
// se il campo e' vuoto restituisce null, per svuotare esplicitamente la colonna.
export async function resolvePersonFieldId(
  peopleService: PeopleService,
  values: FormValues,
  fieldName: string,
  // null (non undefined) e' richiesto dalla REST API di SharePoint per svuotare esplicitamente il campo persona.
  // eslint-disable-next-line @rushstack/no-new-null
): Promise<number | null> {
  const raw = values[fieldName];
  if (raw && typeof raw === 'object') {
    const person = raw as IPersonValue;
    if (person.id !== undefined && person.id !== null) {
      return person.id;
    }
    if (person.loginName) {
      const resolved = await peopleService.ensureUser(person.loginName);
      return resolved.id;
    }
  }
  return null;
}

export function projectToFormValues(project: IProject): FormValues {
  return {
    ProjectCode: project.ProjectCode,
    Title: project.Title,
    Description: project.Description,
    Sponsor: project.Sponsor,
    SponsorUser: personRefToValue(project.SponsorUser),
    SponsorUserEmail: project.SponsorUserEmail,
    ProjectManager: project.ProjectManager,
    ProjectManagerUser: personRefToValue(project.ProjectManagerUser),
    ProjectManagerUserEmail: project.ProjectManagerUserEmail,
    StartDate: project.StartDate,
    EndDate: project.EndDate,
    Status: project.Status,
    Priority: project.Priority,
    RAG: project.RAG,
    Progress: project.Progress,
    BudgetTotal: project.BudgetTotal,
    BudgetCommitted: project.BudgetCommitted,
    BudgetConsumed: project.BudgetConsumed,
    StrategicArea: project.StrategicArea,
    Notes: project.Notes,
  };
}

export function formValuesToNewProject(values: FormValues): NewProject {
  return {
    ProjectCode: typeof values.ProjectCode === 'string' ? values.ProjectCode : '',
    Title: typeof values.Title === 'string' ? values.Title : '',
    Description: values.Description as string | undefined,
    Sponsor: values.Sponsor as string | undefined,
    SponsorUserEmail: values.SponsorUserEmail as string | undefined,
    ProjectManager: values.ProjectManager as string | undefined,
    ProjectManagerUserEmail: values.ProjectManagerUserEmail as string | undefined,
    StartDate: values.StartDate as string | undefined,
    EndDate: values.EndDate as string | undefined,
    Status: values.Status as ProjectStatus,
    Priority: values.Priority as ProjectPriority,
    RAG: values.RAG as ProjectRag,
    Progress: values.Progress as number | undefined,
    BudgetTotal: values.BudgetTotal as number | undefined,
    BudgetCommitted: values.BudgetCommitted as number | undefined,
    BudgetConsumed: values.BudgetConsumed as number | undefined,
    StrategicArea: values.StrategicArea as string | undefined,
    Notes: values.Notes as string | undefined,
    // SponsorUserId/ProjectManagerUserId non sono inclusi qui: la loro
    // risoluzione richiede una chiamata asincrona (ensureUser), fatta dal
    // chiamante con resolvePersonFieldId prima di scrivere.
  };
}

export function deliverableToFormValues(item: IDeliverable): FormValues {
  return {
    ProjectCode: item.ProjectCode,
    Title: item.Title,
    Owner: item.Owner,
    OwnerUser: personRefToValue(item.OwnerUser),
    OwnerUserEmail: item.OwnerUserEmail,
    StartDate: item.StartDate,
    EndDate: item.EndDate,
    Progress: item.Progress,
    Status: item.Status,
    Weight: item.Weight,
    Notes: item.Notes,
  };
}

export function formValuesToNewDeliverable(values: FormValues, projectCode: string): NewDeliverable {
  return {
    ProjectCode: projectCode,
    Title: typeof values.Title === 'string' ? values.Title : '',
    Owner: values.Owner as string | undefined,
    OwnerUserEmail: values.OwnerUserEmail as string | undefined,
    StartDate: values.StartDate as string | undefined,
    EndDate: values.EndDate as string | undefined,
    Progress: values.Progress as number | undefined,
    Status: values.Status as DeliverableStatus,
    Weight: values.Weight as number | undefined,
    Notes: values.Notes as string | undefined,
  };
}

export function issueToFormValues(item: IIssue): FormValues {
  return {
    ProjectCode: item.ProjectCode,
    Title: item.Title,
    Description: item.Description,
    Severity: item.Severity,
    Owner: item.Owner,
    OwnerUser: personRefToValue(item.OwnerUser),
    OwnerUserEmail: item.OwnerUserEmail,
    OpenDate: item.OpenDate,
    DueDate: item.DueDate,
    Status: item.Status,
    Action: item.Action,
    EscalationRequired: item.EscalationRequired,
    Notes: item.Notes,
  };
}

export function formValuesToNewIssue(values: FormValues, projectCode: string): NewIssue {
  return {
    ProjectCode: projectCode,
    Title: typeof values.Title === 'string' ? values.Title : '',
    Description: values.Description as string | undefined,
    Severity: values.Severity as IssueSeverity,
    Owner: values.Owner as string | undefined,
    OwnerUserEmail: values.OwnerUserEmail as string | undefined,
    OpenDate: values.OpenDate as string | undefined,
    DueDate: values.DueDate as string | undefined,
    Status: values.Status as IssueStatus,
    Action: values.Action as string | undefined,
    EscalationRequired: values.EscalationRequired as boolean | undefined,
    Notes: values.Notes as string | undefined,
  };
}

export function costToFormValues(item: ICost): FormValues {
  return {
    ProjectCode: item.ProjectCode,
    CostCategory: item.CostCategory,
    AmountAllocated: item.AmountAllocated,
    AmountCommitted: item.AmountCommitted,
    AmountConsumed: item.AmountConsumed,
    Supplier: item.Supplier,
    Notes: item.Notes,
  };
}

// Come in ExcelService.ts: Costs non ha una colonna Title propria, viene generata
// da ProjectCode+CostCategory (il Title della lista SharePoint esiste comunque di
// default e va valorizzato).
export function formValuesToNewCost(values: FormValues, projectCode: string): NewCost {
  const costCategory = values.CostCategory as CostCategory;
  return {
    ProjectCode: projectCode,
    Title: `${projectCode}-${costCategory}`,
    CostCategory: costCategory,
    AmountAllocated: values.AmountAllocated as number | undefined,
    AmountCommitted: values.AmountCommitted as number | undefined,
    AmountConsumed: values.AmountConsumed as number | undefined,
    Supplier: values.Supplier as string | undefined,
    Notes: values.Notes as string | undefined,
  };
}

export function allocationToFormValues(item: IAllocation): FormValues {
  return {
    ProjectCode: item.ProjectCode,
    ResourceCode: item.ResourceCode,
    AllocationPercent: item.AllocationPercent,
    RoleOnProject: item.RoleOnProject,
    StartDate: item.StartDate,
    EndDate: item.EndDate,
    Notes: item.Notes,
  };
}

// Come in ExcelService.ts: Allocations non ha una colonna Title propria, viene
// generata da ProjectCode+ResourceCode.
export function formValuesToNewAllocation(values: FormValues, projectCode: string): NewAllocation {
  const resourceCode = typeof values.ResourceCode === 'string' ? values.ResourceCode : '';
  return {
    ProjectCode: projectCode,
    Title: `${projectCode}-${resourceCode}`,
    ResourceCode: resourceCode,
    AllocationPercent: values.AllocationPercent as number | undefined,
    RoleOnProject: values.RoleOnProject as string | undefined,
    StartDate: values.StartDate as string | undefined,
    EndDate: values.EndDate as string | undefined,
    Notes: values.Notes as string | undefined,
  };
}

export function resourceToFormValues(item: IResource): FormValues {
  return {
    ResourceCode: item.ResourceCode,
    Title: item.Title,
    PersonUser: personRefToValue(item.PersonUser),
    PersonUserEmail: item.PersonUserEmail,
    Role: item.Role,
    Unit: item.Unit,
    Capacity: item.Capacity,
    Active: item.Active,
    Notes: item.Notes,
  };
}

export function formValuesToNewResource(values: FormValues): NewResource {
  return {
    ResourceCode: typeof values.ResourceCode === 'string' ? values.ResourceCode : '',
    Title: typeof values.Title === 'string' ? values.Title : '',
    PersonUserEmail: values.PersonUserEmail as string | undefined,
    Role: values.Role as string | undefined,
    Unit: values.Unit as string | undefined,
    Capacity: values.Capacity as number | undefined,
    Active: values.Active as boolean | undefined,
    Notes: values.Notes as string | undefined,
  };
}
