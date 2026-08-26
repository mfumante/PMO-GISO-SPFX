import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  Alert,
  Box,
  Button,
  Chip,
  Divider,
  FormControl,
  IconButton,
  InputLabel,
  LinearProgress,
  MenuItem,
  Paper,
  Select,
  Skeleton,
  Snackbar,
  Stack,
  Tab,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Tabs,
  Tooltip,
  Typography,
} from '@mui/material';
import type { SelectChangeEvent } from '@mui/material';
import AddOutlinedIcon from '@mui/icons-material/AddOutlined';
import EditOutlinedIcon from '@mui/icons-material/EditOutlined';
import DeleteOutlinedIcon from '@mui/icons-material/DeleteOutlined';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { PageContainer } from '../layout/PageContainer';
import { useAsyncData } from '../../hooks/useAsyncData';
import { useFeedback } from '../../hooks/useFeedback';
import { EntityFormDialog, FormValues, IFormFieldOption } from '../common/EntityFormDialog';
import { ConfirmDialog } from '../common/ConfirmDialog';
import { PersonDisplay } from '../common/PersonDisplay';
import { KpiCard } from '../common/KpiCard';
import { ExpandableCellText } from '../common/ExpandableCellText';
import { getTableSx, TABLE_MAX_HEIGHT } from '../common/tableStyles';
import { useLayoutPreferences } from '../layout/LayoutPreferencesContext';
import { HelpLabel } from '../help/HelpLabel';
import { getFieldHelp } from '../help/fieldHelp';
import {
  buildAllocationFields,
  buildCostFields,
  buildDeliverableFields,
  buildIssueFields,
  buildProjectFields,
} from '../forms/entityFormSchemas';
import {
  allocationToFormValues,
  costToFormValues,
  deliverableToFormValues,
  formValuesToNewAllocation,
  formValuesToNewCost,
  formValuesToNewDeliverable,
  formValuesToNewIssue,
  formValuesToNewProject,
  issueToFormValues,
  projectToFormValues,
  resolvePersonFieldId,
} from '../forms/entityFormValues';
import { PeopleService } from '../../services/PeopleService';
import {
  DeliverableStatus,
  IAllocation,
  ICost,
  IDeliverable,
  IIssue,
  IProject,
  IResource,
  IssueSeverity,
  SharePointDataService,
} from '../../services/SharePointDataService';

interface SchedaProgettoProps {
  context: WebPartContext;
  projectCode: string | undefined;
  onProjectCodeChange: (projectCode: string) => void;
}

interface IProjectDetailData {
  project: IProject | undefined;
  deliverables: IDeliverable[];
  issues: IIssue[];
  costs: ICost[];
  allocations: IAllocation[];
  resources: IResource[];
}

type DeliverableDialogState = { mode: 'new' } | { mode: 'edit'; item: IDeliverable };
type IssueDialogState = { mode: 'new' } | { mode: 'edit'; item: IIssue };
type CostDialogState = { mode: 'new' } | { mode: 'edit'; item: ICost };
type AllocationDialogState = { mode: 'new' } | { mode: 'edit'; item: IAllocation };

type DeleteTarget =
  | { kind: 'deliverable'; item: IDeliverable }
  | { kind: 'issue'; item: IIssue }
  | { kind: 'cost'; item: ICost }
  | { kind: 'allocation'; item: IAllocation };

const currencyFormatter = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });

const severityChipProps: Record<IssueSeverity, { color: 'error' | 'warning' | 'info' | 'default' }> = {
  Critical: { color: 'error' },
  High: { color: 'warning' },
  Medium: { color: 'info' },
  Low: { color: 'default' },
};

const deliverableStatusChipProps: Record<DeliverableStatus, { color: 'default' | 'info' | 'success' | 'error' }> = {
  'Not Started': { color: 'default' },
  'In Progress': { color: 'info' },
  Completed: { color: 'success' },
  Delayed: { color: 'error' },
  Cancelled: { color: 'default' },
};

function formatDate(value: string | undefined): string {
  if (!value) {
    return '-';
  }
  const date = new Date(value);
  // isNaN globale (non Number.isNaN): la lib TS di questo progetto esclude es2015.core.
  return isNaN(date.getTime()) ? '-' : date.toLocaleDateString('it-IT');
}

function formatCurrency(value: number | undefined): string {
  return value != null ? currencyFormatter.format(value) : '-';
}

// Riga etichetta/valore dei dati anagrafici: usata nella griglia a due colonne
// dell'Overview, cosi' su contenitori medi/ampi le informazioni si dispongono
// affiancate invece che tutte impilate in verticale (meno scroll).
function InfoRow({ label, children }: { label: string; children: React.ReactNode }): React.ReactElement {
  return (
    <Stack direction="row" spacing={2} flexWrap="wrap">
      <Typography variant="body2" color="text.secondary" sx={{ minWidth: 160 }}>
        {label}
      </Typography>
      {children}
    </Stack>
  );
}

function OverviewTab({ project, onEditProject }: { project: IProject; onEditProject: () => void }): React.ReactElement {
  const { widthTier } = useLayoutPreferences();
  const budgetTotal = project.BudgetTotal ?? 0;
  const budgetConsumed = project.BudgetConsumed ?? 0;
  const budgetPercent = budgetTotal > 0 ? Math.min(100, (budgetConsumed / budgetTotal) * 100) : 0;

  const daysRemaining = useMemo(() => {
    if (!project.EndDate) {
      return undefined;
    }
    const end = new Date(project.EndDate).getTime();
    if (isNaN(end)) {
      return undefined;
    }
    return Math.ceil((end - Date.now()) / (1000 * 60 * 60 * 24));
  }, [project.EndDate]);

  return (
    <Stack spacing={2}>
      <Box sx={{ display: 'grid', gridTemplateColumns: widthTier === 'narrow' ? '1fr' : 'repeat(3, 1fr)', gap: 2 }}>
        <KpiCard label="Avanzamento" value={`${Math.round(project.Progress ?? 0)}%`} />
        <KpiCard
          label="Budget consumato / totale"
          value={`${formatCurrency(budgetConsumed)} / ${formatCurrency(budgetTotal)}`}
          helper={
            <Box sx={{ mt: 0.5 }}>
              <LinearProgress variant="determinate" value={budgetPercent} />
            </Box>
          }
        />
        <KpiCard
          label="Giorni residui"
          value={
            daysRemaining == null
              ? '-'
              : daysRemaining >= 0
                ? `${daysRemaining} giorni`
                : `Scaduto da ${Math.abs(daysRemaining)} giorni`
          }
        />
      </Box>

      <Paper variant="outlined" sx={{ p: 2 }}>
        <Stack direction="row" justifyContent="space-between" alignItems="flex-start" spacing={2}>
          <Box sx={{ flexGrow: 1 }}>
            <Box
              sx={{
                display: 'grid',
                gridTemplateColumns: widthTier === 'narrow' ? '1fr' : '1fr 1fr',
                gap: 1.5,
                mb: 1.5,
              }}
            >
              <InfoRow label="Sponsor">
                <PersonDisplay person={project.SponsorUser} fallbackText={project.Sponsor} />
              </InfoRow>
              <InfoRow label="Project Manager">
                <PersonDisplay person={project.ProjectManagerUser} fallbackText={project.ProjectManager} />
              </InfoRow>
              <InfoRow label="Periodo">
                <Typography variant="body2">
                  {formatDate(project.StartDate)} &rarr; {formatDate(project.EndDate)}
                </Typography>
              </InfoRow>
              <InfoRow label="Area strategica">
                <Typography variant="body2">{project.StrategicArea ?? '-'}</Typography>
              </InfoRow>
            </Box>
            <Divider sx={{ mb: 1.5 }} />
            <Box
              sx={{
                display: 'grid',
                gridTemplateColumns: widthTier === 'wide' ? '1fr 1fr' : '1fr',
                gap: 2,
              }}
            >
              <Box>
                <Typography variant="body2" color="text.secondary">
                  Descrizione
                </Typography>
                <Typography variant="body2">{project.Description || 'Nessuna descrizione.'}</Typography>
              </Box>
              {project.Notes && (
                <Box>
                  <Typography variant="body2" color="text.secondary">
                    Note
                  </Typography>
                  <Typography variant="body2">{project.Notes}</Typography>
                </Box>
              )}
            </Box>
          </Box>
          <Button variant="outlined" size="small" startIcon={<EditOutlinedIcon fontSize="small" />} onClick={onEditProject}>
            Modifica anagrafica
          </Button>
        </Stack>
      </Paper>
    </Stack>
  );
}

interface DeliverablesTabProps {
  deliverables: IDeliverable[];
  onAdd: () => void;
  onEdit: (item: IDeliverable) => void;
  onDelete: (item: IDeliverable) => void;
}

function DeliverablesTab({ deliverables, onAdd, onEdit, onDelete }: DeliverablesTabProps): React.ReactElement {
  const { density } = useLayoutPreferences();
  return (
    <Stack spacing={2}>
      <Box>
        <Button variant="contained" size="small" startIcon={<AddOutlinedIcon />} onClick={onAdd}>
          Nuovo deliverable
        </Button>
      </Box>
      {deliverables.length === 0 ? (
        <Alert severity="info">Nessun deliverable registrato per questo progetto.</Alert>
      ) : (
        <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
          <Table size="small" sx={getTableSx(density, 760)}>
            <TableHead>
              <TableRow>
                <TableCell sx={{ width: '22%' }}>Title</TableCell>
                <TableCell sx={{ width: '16%' }}>Owner</TableCell>
                <TableCell sx={{ width: '11%' }}>StartDate</TableCell>
                <TableCell sx={{ width: '11%' }}>EndDate</TableCell>
                <TableCell sx={{ width: '16%' }}>
                  <HelpLabel label="Progress" help={getFieldHelp('Deliverable', 'Progress')} />
                </TableCell>
                <TableCell sx={{ width: '12%' }}>Status</TableCell>
                <TableCell align="right" sx={{ width: '6%' }}>
                  <HelpLabel label="Weight" help={getFieldHelp('Deliverable', 'Weight')} />
                </TableCell>
                <TableCell align="right" sx={{ width: '6%' }}>Azioni</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {deliverables.map((deliverable) => (
                <TableRow key={deliverable.Id}>
                  <TableCell>
                    <ExpandableCellText text={deliverable.Title} />
                  </TableCell>
                  <TableCell>
                    <PersonDisplay person={deliverable.OwnerUser} fallbackText={deliverable.Owner} />
                  </TableCell>
                  <TableCell>{formatDate(deliverable.StartDate)}</TableCell>
                  <TableCell>{formatDate(deliverable.EndDate)}</TableCell>
                  <TableCell>
                    <Stack direction="row" spacing={1} alignItems="center">
                      <Box sx={{ flexGrow: 1 }}>
                        <LinearProgress variant="determinate" value={Math.min(100, Math.max(0, deliverable.Progress ?? 0))} />
                      </Box>
                      <Typography variant="caption" color="text.secondary">
                        {Math.round(deliverable.Progress ?? 0)}%
                      </Typography>
                    </Stack>
                  </TableCell>
                  <TableCell>
                    <Chip size="small" label={deliverable.Status} color={deliverableStatusChipProps[deliverable.Status].color} />
                  </TableCell>
                  <TableCell align="right">{deliverable.Weight ?? '-'}</TableCell>
                  <TableCell align="right">
                    <Tooltip title="Modifica">
                      <IconButton size="small" onClick={() => onEdit(deliverable)}>
                        <EditOutlinedIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                    <Tooltip title="Elimina">
                      <IconButton size="small" onClick={() => onDelete(deliverable)}>
                        <DeleteOutlinedIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      )}
    </Stack>
  );
}

interface IssuesTabProps {
  issues: IIssue[];
  onAdd: () => void;
  onEdit: (item: IIssue) => void;
  onDelete: (item: IIssue) => void;
}

function IssuesTab({ issues, onAdd, onEdit, onDelete }: IssuesTabProps): React.ReactElement {
  const { density } = useLayoutPreferences();
  return (
    <Stack spacing={2}>
      <Box>
        <Button variant="contained" size="small" startIcon={<AddOutlinedIcon />} onClick={onAdd}>
          Nuova issue
        </Button>
      </Box>
      {issues.length === 0 ? (
        <Alert severity="info">Nessuna issue registrata per questo progetto.</Alert>
      ) : (
        <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
          <Table size="small" sx={getTableSx(density, 760)}>
            <TableHead>
              <TableRow>
                <TableCell sx={{ width: '22%' }}>Title</TableCell>
                <TableCell sx={{ width: '11%' }}>Severity</TableCell>
                <TableCell sx={{ width: '16%' }}>Owner</TableCell>
                <TableCell sx={{ width: '11%' }}>DueDate</TableCell>
                <TableCell sx={{ width: '22%' }}>Action</TableCell>
                <TableCell sx={{ width: '10%' }}>
                  <HelpLabel label="Escalation" help={getFieldHelp('Issue', 'EscalationRequired')} />
                </TableCell>
                <TableCell align="right" sx={{ width: '8%' }}>Azioni</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {issues.map((issue) => (
                <TableRow key={issue.Id}>
                  <TableCell>
                    <ExpandableCellText text={issue.Title} />
                  </TableCell>
                  <TableCell>
                    <Chip size="small" label={issue.Severity} color={severityChipProps[issue.Severity].color} />
                  </TableCell>
                  <TableCell>
                    <PersonDisplay person={issue.OwnerUser} fallbackText={issue.Owner} />
                  </TableCell>
                  <TableCell>{formatDate(issue.DueDate)}</TableCell>
                  <TableCell>
                    <ExpandableCellText text={issue.Action} />
                  </TableCell>
                  <TableCell>
                    <Chip
                      size="small"
                      label={issue.EscalationRequired ? 'Si' : 'No'}
                      color={issue.EscalationRequired ? 'warning' : 'default'}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <Tooltip title="Modifica">
                      <IconButton size="small" onClick={() => onEdit(issue)}>
                        <EditOutlinedIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                    <Tooltip title="Elimina">
                      <IconButton size="small" onClick={() => onDelete(issue)}>
                        <DeleteOutlinedIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      )}
    </Stack>
  );
}

interface CostsTabProps {
  costs: ICost[];
  onAdd: () => void;
  onEdit: (item: ICost) => void;
  onDelete: (item: ICost) => void;
}

function CostsTab({ costs, onAdd, onEdit, onDelete }: CostsTabProps): React.ReactElement {
  const { density } = useLayoutPreferences();
  const totals = costs.reduce(
    (acc, cost) => ({
      allocated: acc.allocated + (cost.AmountAllocated ?? 0),
      committed: acc.committed + (cost.AmountCommitted ?? 0),
      consumed: acc.consumed + (cost.AmountConsumed ?? 0),
    }),
    { allocated: 0, committed: 0, consumed: 0 },
  );

  return (
    <Stack spacing={2}>
      <Box>
        <Button variant="contained" size="small" startIcon={<AddOutlinedIcon />} onClick={onAdd}>
          Nuovo costo
        </Button>
      </Box>
      {costs.length === 0 ? (
        <Alert severity="info">Nessun costo registrato per questo progetto.</Alert>
      ) : (
        <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
          <Table size="small" sx={getTableSx(density, 640)}>
            <TableHead>
              <TableRow>
                <TableCell sx={{ width: '24%' }}>CostCategory</TableCell>
                <TableCell align="right" sx={{ width: '16%' }}>Allocated</TableCell>
                <TableCell align="right" sx={{ width: '16%' }}>Committed</TableCell>
                <TableCell align="right" sx={{ width: '16%' }}>Consumed</TableCell>
                <TableCell sx={{ width: '18%' }}>Supplier</TableCell>
                <TableCell align="right" sx={{ width: '10%' }}>Azioni</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {costs.map((cost) => (
                <TableRow key={cost.Id}>
                  <TableCell>{cost.CostCategory}</TableCell>
                  <TableCell align="right">{formatCurrency(cost.AmountAllocated)}</TableCell>
                  <TableCell align="right">{formatCurrency(cost.AmountCommitted)}</TableCell>
                  <TableCell align="right">{formatCurrency(cost.AmountConsumed)}</TableCell>
                  <TableCell>
                    <ExpandableCellText text={cost.Supplier} />
                  </TableCell>
                  <TableCell align="right">
                    <Tooltip title="Modifica">
                      <IconButton size="small" onClick={() => onEdit(cost)}>
                        <EditOutlinedIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                    <Tooltip title="Elimina">
                      <IconButton size="small" onClick={() => onDelete(cost)}>
                        <DeleteOutlinedIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              ))}
              <TableRow sx={{ '& td': { fontWeight: 700, borderTop: '2px solid', borderTopColor: 'divider' } }}>
                <TableCell>Totale</TableCell>
                <TableCell align="right">{currencyFormatter.format(totals.allocated)}</TableCell>
                <TableCell align="right">{currencyFormatter.format(totals.committed)}</TableCell>
                <TableCell align="right">{currencyFormatter.format(totals.consumed)}</TableCell>
                <TableCell />
                <TableCell />
              </TableRow>
            </TableBody>
          </Table>
        </TableContainer>
      )}
    </Stack>
  );
}

interface ResourcesTabProps {
  allocations: IAllocation[];
  resources: IResource[];
  onAdd: () => void;
  onEdit: (item: IAllocation) => void;
  onDelete: (item: IAllocation) => void;
}

function ResourcesTab({ allocations, resources, onAdd, onEdit, onDelete }: ResourcesTabProps): React.ReactElement {
  const { density } = useLayoutPreferences();
  const resourceByCode = new Map(resources.map((resource): [string, IResource] => [resource.ResourceCode, resource]));

  return (
    <Stack spacing={2}>
      <Box>
        <Button variant="contained" size="small" startIcon={<AddOutlinedIcon />} onClick={onAdd} disabled={resources.length === 0}>
          Nuova allocazione
        </Button>
        {resources.length === 0 && (
          <Typography variant="caption" color="text.secondary" sx={{ ml: 2 }}>
            Nessuna risorsa disponibile in PMO_Resources.
          </Typography>
        )}
      </Box>
      {allocations.length === 0 ? (
        <Alert severity="info">Nessuna risorsa allocata su questo progetto.</Alert>
      ) : (
        <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
          <Table size="small" sx={getTableSx(density, 700)}>
            <TableHead>
              <TableRow>
                <TableCell sx={{ width: '14%' }}>ResourceCode</TableCell>
                <TableCell sx={{ width: '26%' }}>Risorsa</TableCell>
                <TableCell sx={{ width: '18%' }}>RoleOnProject</TableCell>
                <TableCell align="right" sx={{ width: '12%' }}>
                  <HelpLabel label="Allocazione" help={getFieldHelp('Allocation', 'AllocationPercent')} />
                </TableCell>
                <TableCell sx={{ width: '20%' }}>Periodo</TableCell>
                <TableCell align="right" sx={{ width: '10%' }}>Azioni</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {allocations.map((allocation) => {
                const resource = resourceByCode.get(allocation.ResourceCode);
                const resourceLabel = resource ? `${resource.Title} (${resource.Role ?? '-'})` : '-';
                return (
                  <TableRow key={allocation.Id}>
                    <TableCell sx={{ fontFamily: 'monospace' }}>{allocation.ResourceCode}</TableCell>
                    <TableCell>
                      <ExpandableCellText text={resourceLabel} />
                    </TableCell>
                    <TableCell>
                      <ExpandableCellText text={allocation.RoleOnProject} />
                    </TableCell>
                    <TableCell align="right">{allocation.AllocationPercent ?? 0}%</TableCell>
                    <TableCell>
                      {formatDate(allocation.StartDate)} &rarr; {formatDate(allocation.EndDate)}
                    </TableCell>
                    <TableCell align="right">
                      <Tooltip title="Modifica">
                        <IconButton size="small" onClick={() => onEdit(allocation)}>
                          <EditOutlinedIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="Elimina">
                        <IconButton size="small" onClick={() => onDelete(allocation)}>
                          <DeleteOutlinedIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                    </TableCell>
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
        </TableContainer>
      )}
    </Stack>
  );
}

function getDeleteMessage(target: DeleteTarget): React.ReactNode {
  switch (target.kind) {
    case 'deliverable':
      return (
        <>Confermi l&apos;eliminazione del deliverable <strong>{target.item.Title}</strong>?</>
      );
    case 'issue':
      return (
        <>Confermi l&apos;eliminazione della issue <strong>{target.item.Title}</strong>?</>
      );
    case 'cost':
      return (
        <>Confermi l&apos;eliminazione del costo <strong>{target.item.CostCategory}</strong>?</>
      );
    case 'allocation':
      return (
        <>Confermi l&apos;eliminazione dell&apos;allocazione sulla risorsa <strong>{target.item.ResourceCode}</strong>?</>
      );
    default:
      return '';
  }
}

export default function SchedaProgetto({ context, projectCode, onProjectCodeChange }: SchedaProgettoProps): React.ReactElement {
  const service = useMemo(() => new SharePointDataService(context), [context]);
  const peopleService = useMemo(() => new PeopleService(context), [context]);
  const { feedback, showSuccess, showError, closeFeedback } = useFeedback();
  const [activeTab, setActiveTab] = useState(0);
  const [refreshKey, setRefreshKey] = useState(0);

  const [projectEditOpen, setProjectEditOpen] = useState(false);
  const [deliverableDialog, setDeliverableDialog] = useState<DeliverableDialogState | null>(null);
  const [issueDialog, setIssueDialog] = useState<IssueDialogState | null>(null);
  const [costDialog, setCostDialog] = useState<CostDialogState | null>(null);
  const [allocationDialog, setAllocationDialog] = useState<AllocationDialogState | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<DeleteTarget | null>(null);
  const [deleting, setDeleting] = useState(false);

  const projectsResult = useAsyncData(() => service.getProjects(), [service, refreshKey]);

  // Se non c'e' ancora un progetto selezionato, seleziona il primo disponibile.
  useEffect(() => {
    if (!projectCode && projectsResult.status === 'ready' && projectsResult.data.length > 0) {
      onProjectCodeChange(projectsResult.data[0].ProjectCode);
    }
  }, [projectCode, projectsResult, onProjectCodeChange]);

  const detailResult = useAsyncData<IProjectDetailData>(async () => {
    if (!projectCode) {
      return { project: undefined, deliverables: [], issues: [], costs: [], allocations: [], resources: [] };
    }
    const [project, deliverables, issues, costs, allocations, resources] = await Promise.all([
      service.getProjectByCode(projectCode),
      service.getDeliverablesByProject(projectCode),
      service.getIssuesByProject(projectCode),
      service.getCostsByProject(projectCode),
      service.getAllocationsByProject(projectCode),
      service.getResources(),
    ]);
    return { project, deliverables, issues, costs, allocations, resources };
  }, [service, projectCode, refreshKey]);

  function handleProjectSelectChange(event: SelectChangeEvent): void {
    onProjectCodeChange(event.target.value);
  }

  function refreshAfterWrite(): void {
    setRefreshKey((key) => key + 1);
  }

  async function handleSaveProjectEdit(values: FormValues): Promise<void> {
    const project = detailResult.status === 'ready' ? detailResult.data.project : undefined;
    if (!project) {
      return;
    }
    const payload = formValuesToNewProject(values);
    payload.SponsorUserId = await resolvePersonFieldId(peopleService, values, 'SponsorUser');
    payload.ProjectManagerUserId = await resolvePersonFieldId(peopleService, values, 'ProjectManagerUser');
    const writeResult = await service.updateProject(project.Id, payload);
    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }
    refreshAfterWrite();
    showSuccess('Progetto aggiornato.');
    setProjectEditOpen(false);
  }

  async function handleSaveDeliverable(values: FormValues): Promise<void> {
    if (!projectCode) {
      return;
    }
    const payload = formValuesToNewDeliverable(values, projectCode);
    payload.OwnerUserId = await resolvePersonFieldId(peopleService, values, 'OwnerUser');
    const target = deliverableDialog;
    const writeResult =
      target?.mode === 'edit' ? await service.updateDeliverable(target.item.Id, payload) : await service.createDeliverable(payload);
    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }
    refreshAfterWrite();
    showSuccess(target?.mode === 'edit' ? 'Deliverable aggiornato.' : 'Deliverable creato.');
    setDeliverableDialog(null);
  }

  async function handleSaveIssue(values: FormValues): Promise<void> {
    if (!projectCode) {
      return;
    }
    const payload = formValuesToNewIssue(values, projectCode);
    payload.OwnerUserId = await resolvePersonFieldId(peopleService, values, 'OwnerUser');
    const target = issueDialog;
    const writeResult = target?.mode === 'edit' ? await service.updateIssue(target.item.Id, payload) : await service.createIssue(payload);
    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }
    refreshAfterWrite();
    showSuccess(target?.mode === 'edit' ? 'Issue aggiornata.' : 'Issue creata.');
    setIssueDialog(null);
  }

  async function handleSaveCost(values: FormValues): Promise<void> {
    if (!projectCode) {
      return;
    }
    const payload = formValuesToNewCost(values, projectCode);
    const target = costDialog;
    const writeResult = target?.mode === 'edit' ? await service.updateCost(target.item.Id, payload) : await service.createCost(payload);
    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }
    refreshAfterWrite();
    showSuccess(target?.mode === 'edit' ? 'Costo aggiornato.' : 'Costo creato.');
    setCostDialog(null);
  }

  async function handleSaveAllocation(values: FormValues): Promise<void> {
    if (!projectCode) {
      return;
    }
    const payload = formValuesToNewAllocation(values, projectCode);
    const target = allocationDialog;
    const writeResult =
      target?.mode === 'edit' ? await service.updateAllocation(target.item.Id, payload) : await service.createAllocation(payload);
    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }
    refreshAfterWrite();
    showSuccess(target?.mode === 'edit' ? 'Allocazione aggiornata.' : 'Allocazione creata.');
    setAllocationDialog(null);
  }

  async function handleConfirmDelete(): Promise<void> {
    if (!deleteTarget) {
      return;
    }
    setDeleting(true);
    try {
      const writeResult =
        deleteTarget.kind === 'deliverable'
          ? await service.deleteDeliverable(deleteTarget.item.Id)
          : deleteTarget.kind === 'issue'
            ? await service.deleteIssue(deleteTarget.item.Id)
            : deleteTarget.kind === 'cost'
              ? await service.deleteCost(deleteTarget.item.Id)
              : await service.deleteAllocation(deleteTarget.item.Id);

      if (!writeResult.success) {
        showError(writeResult.message);
        return;
      }
      refreshAfterWrite();
      showSuccess('Elemento eliminato.');
      setDeleteTarget(null);
    } catch (error) {
      showError(error instanceof Error ? error.message : 'Errore durante l\'eliminazione.');
    } finally {
      setDeleting(false);
    }
  }

  const resourceOptions: IFormFieldOption[] =
    detailResult.status === 'ready'
      ? detailResult.data.resources.map((resource) => ({ value: resource.ResourceCode, label: `${resource.ResourceCode} - ${resource.Title}` }))
      : [];

  return (
    <PageContainer title="Scheda Progetto">
      <Stack spacing={3}>
        <FormControl size="small" sx={{ minWidth: 320 }}>
          <InputLabel id="project-select-label">Progetto</InputLabel>
          <Select
            labelId="project-select-label"
            label="Progetto"
            value={projectsResult.status === 'ready' ? (projectCode ?? '') : ''}
            onChange={handleProjectSelectChange}
            disabled={projectsResult.status !== 'ready' || projectsResult.data.length === 0}
          >
            {projectsResult.status === 'ready' &&
              projectsResult.data.map((project) => (
                <MenuItem key={project.ProjectCode} value={project.ProjectCode}>
                  {project.ProjectCode} &ndash; {project.Title}
                </MenuItem>
              ))}
          </Select>
        </FormControl>

        {projectsResult.status === 'loading' && <Skeleton variant="rectangular" height={40} />}
        {projectsResult.status === 'missing' && (
          <Alert severity="warning">
            La lista <strong>{projectsResult.listTitle}</strong> non esiste ancora. Vai in{' '}
            <strong>Amministrazione</strong> per inizializzarla.
          </Alert>
        )}
        {projectsResult.status === 'error' && <Alert severity="error">{projectsResult.message}</Alert>}
        {projectsResult.status === 'ready' && projectsResult.data.length === 0 && (
          <Alert severity="info">
            Non ci sono ancora progetti in <strong>PMO_Projects</strong>. Vai in <strong>Amministrazione</strong> per
            importare i dati, oppure creane uno dal Portfolio Progetti.
          </Alert>
        )}

        {projectsResult.status === 'ready' && projectsResult.data.length > 0 && (
          <>
            {detailResult.status === 'loading' && (
              <Stack spacing={1}>
                <Skeleton variant="rectangular" height={120} />
                <Skeleton variant="rectangular" height={200} />
              </Stack>
            )}
            {detailResult.status === 'missing' && (
              <Alert severity="warning">
                La lista <strong>{detailResult.listTitle}</strong> non esiste ancora. Vai in{' '}
                <strong>Amministrazione</strong> per inizializzarla.
              </Alert>
            )}
            {detailResult.status === 'error' && <Alert severity="error">{detailResult.message}</Alert>}
            {detailResult.status === 'ready' && !detailResult.data.project && (
              <Alert severity="warning">Progetto non trovato.</Alert>
            )}

            {detailResult.status === 'ready' && detailResult.data.project && (
              <>
                <Tabs value={activeTab} onChange={(_event, value) => setActiveTab(value)}>
                  <Tab label="Overview" />
                  <Tab label="Deliverable" />
                  <Tab label="Issue" />
                  <Tab label="Costi" />
                  <Tab label="Risorse" />
                </Tabs>

                <Box>
                  {activeTab === 0 && (
                    <OverviewTab project={detailResult.data.project} onEditProject={() => setProjectEditOpen(true)} />
                  )}
                  {activeTab === 1 && (
                    <DeliverablesTab
                      deliverables={detailResult.data.deliverables}
                      onAdd={() => setDeliverableDialog({ mode: 'new' })}
                      onEdit={(item) => setDeliverableDialog({ mode: 'edit', item })}
                      onDelete={(item) => setDeleteTarget({ kind: 'deliverable', item })}
                    />
                  )}
                  {activeTab === 2 && (
                    <IssuesTab
                      issues={detailResult.data.issues}
                      onAdd={() => setIssueDialog({ mode: 'new' })}
                      onEdit={(item) => setIssueDialog({ mode: 'edit', item })}
                      onDelete={(item) => setDeleteTarget({ kind: 'issue', item })}
                    />
                  )}
                  {activeTab === 3 && (
                    <CostsTab
                      costs={detailResult.data.costs}
                      onAdd={() => setCostDialog({ mode: 'new' })}
                      onEdit={(item) => setCostDialog({ mode: 'edit', item })}
                      onDelete={(item) => setDeleteTarget({ kind: 'cost', item })}
                    />
                  )}
                  {activeTab === 4 && (
                    <ResourcesTab
                      allocations={detailResult.data.allocations}
                      resources={detailResult.data.resources}
                      onAdd={() => setAllocationDialog({ mode: 'new' })}
                      onEdit={(item) => setAllocationDialog({ mode: 'edit', item })}
                      onDelete={(item) => setDeleteTarget({ kind: 'allocation', item })}
                    />
                  )}
                </Box>
              </>
            )}
          </>
        )}
      </Stack>

      {projectEditOpen && detailResult.status === 'ready' && detailResult.data.project && (
        <EntityFormDialog
          context={context}
          open
          title="Modifica anagrafica progetto"
          fields={buildProjectFields(true)}
          initialValues={projectToFormValues(detailResult.data.project)}
          dateRange={{ startField: 'StartDate', endField: 'EndDate' }}
          onCancel={() => setProjectEditOpen(false)}
          onSave={handleSaveProjectEdit}
        />
      )}

      {deliverableDialog && (
        <EntityFormDialog
          context={context}
          open
          title={deliverableDialog.mode === 'edit' ? 'Modifica deliverable' : 'Nuovo deliverable'}
          fields={buildDeliverableFields()}
          initialValues={
            deliverableDialog.mode === 'edit' ? deliverableToFormValues(deliverableDialog.item) : { ProjectCode: projectCode }
          }
          dateRange={{ startField: 'StartDate', endField: 'EndDate' }}
          onCancel={() => setDeliverableDialog(null)}
          onSave={handleSaveDeliverable}
        />
      )}

      {issueDialog && (
        <EntityFormDialog
          context={context}
          open
          title={issueDialog.mode === 'edit' ? 'Modifica issue' : 'Nuova issue'}
          fields={buildIssueFields()}
          initialValues={issueDialog.mode === 'edit' ? issueToFormValues(issueDialog.item) : { ProjectCode: projectCode }}
          dateRange={{ startField: 'OpenDate', endField: 'DueDate' }}
          onCancel={() => setIssueDialog(null)}
          onSave={handleSaveIssue}
        />
      )}

      {costDialog && (
        <EntityFormDialog
          context={context}
          open
          title={costDialog.mode === 'edit' ? 'Modifica costo' : 'Nuovo costo'}
          fields={buildCostFields()}
          initialValues={costDialog.mode === 'edit' ? costToFormValues(costDialog.item) : { ProjectCode: projectCode }}
          onCancel={() => setCostDialog(null)}
          onSave={handleSaveCost}
        />
      )}

      {allocationDialog && (
        <EntityFormDialog
          context={context}
          open
          title={allocationDialog.mode === 'edit' ? 'Modifica allocazione' : 'Nuova allocazione'}
          fields={buildAllocationFields(resourceOptions)}
          initialValues={
            allocationDialog.mode === 'edit' ? allocationToFormValues(allocationDialog.item) : { ProjectCode: projectCode }
          }
          dateRange={{ startField: 'StartDate', endField: 'EndDate' }}
          onCancel={() => setAllocationDialog(null)}
          onSave={handleSaveAllocation}
        />
      )}

      <ConfirmDialog
        open={!!deleteTarget}
        title="Conferma eliminazione"
        message={deleteTarget && getDeleteMessage(deleteTarget)}
        loading={deleting}
        onCancel={() => setDeleteTarget(null)}
        onConfirm={handleConfirmDelete}
      />

      <Snackbar open={feedback.open} autoHideDuration={4000} onClose={closeFeedback} anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}>
        <Alert onClose={closeFeedback} severity={feedback.severity} sx={{ width: '100%' }}>
          {feedback.message}
        </Alert>
      </Snackbar>
    </PageContainer>
  );
}
