import * as React from 'react';
import { useMemo, useState } from 'react';
import {
  Alert,
  Box,
  Button,
  Chip,
  FormControl,
  IconButton,
  InputLabel,
  LinearProgress,
  MenuItem,
  Select,
  Skeleton,
  Snackbar,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TextField,
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
import { EntityFormDialog, FormValues } from '../common/EntityFormDialog';
import { ConfirmDialog } from '../common/ConfirmDialog';
import { PersonDisplay } from '../common/PersonDisplay';
import { ExpandableCellText } from '../common/ExpandableCellText';
import { HelpLabel } from '../help/HelpLabel';
import { getFieldHelp } from '../help/fieldHelp';
import { useLayoutPreferences } from '../layout/LayoutPreferencesContext';
import { getTableSx, TABLE_MAX_HEIGHT } from '../common/tableStyles';
import { buildProjectFields } from '../forms/entityFormSchemas';
import { formValuesToNewProject, projectToFormValues, resolvePersonFieldId } from '../forms/entityFormValues';
import { PeopleService } from '../../services/PeopleService';
import {
  getChoiceOptions,
  IProject,
  ProjectPriority,
  ProjectRag,
  SharePointDataService,
} from '../../services/SharePointDataService';

interface PortfolioProgettiProps {
  context: WebPartContext;
  onOpenProject: (projectCode: string) => void;
}

type ProjectFormTarget = { mode: 'new' } | { mode: 'edit'; project: IProject };

const ALL_VALUE = '__all__';

const currencyFormatter = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });

const ragChipProps: Record<ProjectRag, { color: 'success' | 'warning' | 'error' | 'default' }> = {
  Green: { color: 'success' },
  Amber: { color: 'warning' },
  Red: { color: 'error' },
  Grey: { color: 'default' },
};

const priorityChipProps: Record<ProjectPriority, { color: 'error' | 'warning' | 'info' | 'default' }> = {
  Critical: { color: 'error' },
  High: { color: 'warning' },
  Medium: { color: 'info' },
  Low: { color: 'default' },
};

function matchesSearch(project: IProject, search: string): boolean {
  if (!search) {
    return true;
  }
  const needle = search.trim().toLowerCase();
  return project.Title.toLowerCase().indexOf(needle) !== -1 || project.ProjectCode.toLowerCase().indexOf(needle) !== -1;
}

export default function PortfolioProgetti({ context, onOpenProject }: PortfolioProgettiProps): React.ReactElement {
  const service = useMemo(() => new SharePointDataService(context), [context]);
  const peopleService = useMemo(() => new PeopleService(context), [context]);
  const { feedback, showSuccess, showError, closeFeedback } = useFeedback();
  const { widthTier, density } = useLayoutPreferences();
  const showMediumColumns = widthTier !== 'narrow';
  const showWideColumns = widthTier === 'wide';
  // Larghezza minima della tabella: sotto questa soglia il TableContainer
  // diventa scorrevole lateralmente invece di schiacciare le colonne visibili
  // (vedi getTableSx). Cresce con le colonne aggiuntive delle fasce medie/ampie.
  const tableMinWidth = 640 + (showMediumColumns ? 220 : 0) + (showWideColumns ? 120 : 0);

  const [refreshKey, setRefreshKey] = useState(0);
  const result = useAsyncData(() => service.getProjects(), [service, refreshKey]);

  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState(ALL_VALUE);
  const [priorityFilter, setPriorityFilter] = useState(ALL_VALUE);

  const [formTarget, setFormTarget] = useState<ProjectFormTarget | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<IProject | null>(null);
  const [deleting, setDeleting] = useState(false);

  const statusOptions = useMemo(() => getChoiceOptions('PMO_Projects', 'Status'), []);
  const priorityOptions = useMemo(() => getChoiceOptions('PMO_Projects', 'Priority'), []);

  const filteredProjects = useMemo(() => {
    if (result.status !== 'ready') {
      return [];
    }
    return result.data.filter((project) => {
      if (statusFilter !== ALL_VALUE && project.Status !== statusFilter) {
        return false;
      }
      if (priorityFilter !== ALL_VALUE && project.Priority !== priorityFilter) {
        return false;
      }
      return matchesSearch(project, search);
    });
  }, [result, search, statusFilter, priorityFilter]);

  function handleStatusChange(event: SelectChangeEvent): void {
    setStatusFilter(event.target.value);
  }

  function handlePriorityChange(event: SelectChangeEvent): void {
    setPriorityFilter(event.target.value);
  }

  async function handleSaveProject(values: FormValues): Promise<void> {
    const payload = formValuesToNewProject(values);
    payload.SponsorUserId = await resolvePersonFieldId(peopleService, values, 'SponsorUser');
    payload.ProjectManagerUserId = await resolvePersonFieldId(peopleService, values, 'ProjectManagerUser');

    const target = formTarget;
    const writeResult = target?.mode === 'edit' ? await service.updateProject(target.project.Id, payload) : await service.createProject(payload);

    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }

    setRefreshKey((key) => key + 1);
    showSuccess(target?.mode === 'edit' ? 'Progetto aggiornato.' : 'Progetto creato.');
    setFormTarget(null);
  }

  async function handleConfirmDelete(): Promise<void> {
    if (!deleteTarget) {
      return;
    }
    setDeleting(true);
    try {
      const writeResult = await service.deleteProject(deleteTarget.Id);
      if (!writeResult.success) {
        showError(writeResult.message);
        return;
      }
      setRefreshKey((key) => key + 1);
      showSuccess('Progetto eliminato.');
      setDeleteTarget(null);
    } catch (error) {
      showError(error instanceof Error ? error.message : 'Errore durante l\'eliminazione.');
    } finally {
      setDeleting(false);
    }
  }

  return (
    <PageContainer
      title="Portfolio Progetti"
      actions={
        <Button variant="contained" startIcon={<AddOutlinedIcon />} onClick={() => setFormTarget({ mode: 'new' })}>
          Nuovo progetto
        </Button>
      }
    >
      <Stack spacing={2}>
        <Stack direction="row" spacing={2} flexWrap="wrap" useFlexGap alignItems="center">
          <TextField
            label="Cerca per titolo o codice"
            size="small"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            sx={{ minWidth: 260 }}
          />
          <FormControl size="small" sx={{ minWidth: 180 }}>
            <InputLabel id="status-filter-label">Status</InputLabel>
            <Select labelId="status-filter-label" label="Status" value={statusFilter} onChange={handleStatusChange}>
              <MenuItem value={ALL_VALUE}>Tutti</MenuItem>
              {statusOptions.map((option) => (
                <MenuItem key={option} value={option}>
                  {option}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
          <FormControl size="small" sx={{ minWidth: 180 }}>
            <InputLabel id="priority-filter-label">Priority</InputLabel>
            <Select labelId="priority-filter-label" label="Priority" value={priorityFilter} onChange={handlePriorityChange}>
              <MenuItem value={ALL_VALUE}>Tutte</MenuItem>
              {priorityOptions.map((option) => (
                <MenuItem key={option} value={option}>
                  {option}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </Stack>

        {result.status === 'loading' && (
          <Stack spacing={1}>
            {[1, 2, 3, 4].map((key) => (
              <Skeleton key={key} variant="rectangular" height={48} />
            ))}
          </Stack>
        )}

        {result.status === 'missing' && (
          <Alert severity="warning">
            La lista <strong>{result.listTitle}</strong> non esiste ancora. Vai in <strong>Amministrazione</strong> e
            premi &quot;Inizializza Ambiente&quot; per crearla.
          </Alert>
        )}

        {result.status === 'error' && <Alert severity="error">{result.message}</Alert>}

        {result.status === 'ready' && result.data.length === 0 && (
          <Alert severity="info">
            Non ci sono ancora progetti in <strong>PMO_Projects</strong>. Vai in <strong>Amministrazione</strong> per
            importare i dati, oppure crea il primo progetto con il pulsante &quot;Nuovo progetto&quot;.
          </Alert>
        )}

        {result.status === 'ready' && result.data.length > 0 && (
          <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
            <Table size="small" sx={getTableSx(density, tableMinWidth)}>
              <TableHead>
                <TableRow>
                  <TableCell sx={{ width: '10%' }}>ProjectCode</TableCell>
                  <TableCell sx={{ width: showMediumColumns ? '24%' : '38%' }}>Title</TableCell>
                  {showMediumColumns && <TableCell sx={{ width: '16%' }}>ProjectManager</TableCell>}
                  {showMediumColumns && (
                    <TableCell sx={{ width: '9%' }}>
                      <HelpLabel label="Priority" help={getFieldHelp('Project', 'Priority')} />
                    </TableCell>
                  )}
                  <TableCell sx={{ width: '11%' }}>Status</TableCell>
                  <TableCell sx={{ width: '8%' }}>
                    <HelpLabel label="RAG" help={getFieldHelp('Project', 'RAG')} />
                  </TableCell>
                  <TableCell sx={{ width: showMediumColumns ? '12%' : '18%' }}>
                    <HelpLabel label="Progress" help={getFieldHelp('Project', 'Progress')} />
                  </TableCell>
                  {showWideColumns && <TableCell align="right" sx={{ width: '12%' }}>BudgetTotal</TableCell>}
                  <TableCell
                    align="right"
                    sx={{ width: '9%', position: 'sticky', right: 0, zIndex: 2, backgroundColor: 'background.paper' }}
                  >
                    Azioni
                  </TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {filteredProjects.map((project) => (
                  <TableRow
                    key={project.Id}
                    hover
                    onClick={() => onOpenProject(project.ProjectCode)}
                    sx={{ cursor: 'pointer' }}
                  >
                    <TableCell sx={{ fontFamily: 'monospace' }}>{project.ProjectCode}</TableCell>
                    <TableCell>
                      <ExpandableCellText text={project.Title} stopPropagation />
                    </TableCell>
                    {showMediumColumns && (
                      <TableCell>
                        <PersonDisplay person={project.ProjectManagerUser} fallbackText={project.ProjectManager} />
                      </TableCell>
                    )}
                    {showMediumColumns && (
                      <TableCell>
                        <Chip size="small" label={project.Priority} color={priorityChipProps[project.Priority].color} />
                      </TableCell>
                    )}
                    <TableCell>{project.Status}</TableCell>
                    <TableCell>
                      <Chip size="small" label={project.RAG} color={ragChipProps[project.RAG].color} />
                    </TableCell>
                    <TableCell>
                      <Stack direction="row" spacing={1} alignItems="center">
                        <Box sx={{ flexGrow: 1 }}>
                          <LinearProgress variant="determinate" value={Math.min(100, Math.max(0, project.Progress ?? 0))} />
                        </Box>
                        <Typography variant="caption" color="text.secondary">
                          {Math.round(project.Progress ?? 0)}%
                        </Typography>
                      </Stack>
                    </TableCell>
                    {showWideColumns && (
                      <TableCell align="right">
                        {project.BudgetTotal !== undefined && project.BudgetTotal !== null
                          ? currencyFormatter.format(project.BudgetTotal)
                          : '-'}
                      </TableCell>
                    )}
                    <TableCell
                      align="right"
                      onClick={(event) => event.stopPropagation()}
                      sx={{ position: 'sticky', right: 0, backgroundColor: 'background.paper' }}
                    >
                      <Tooltip title="Modifica">
                        <IconButton size="small" onClick={() => setFormTarget({ mode: 'edit', project })}>
                          <EditOutlinedIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="Elimina">
                        <IconButton size="small" onClick={() => setDeleteTarget(project)}>
                          <DeleteOutlinedIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                    </TableCell>
                  </TableRow>
                ))}
                {filteredProjects.length === 0 && (
                  <TableRow>
                    <TableCell colSpan={6 + (showMediumColumns ? 2 : 0) + (showWideColumns ? 1 : 0)}>
                      <Typography variant="body2" color="text.secondary" sx={{ py: 2, textAlign: 'center' }}>
                        Nessun progetto corrisponde ai filtri selezionati.
                      </Typography>
                    </TableCell>
                  </TableRow>
                )}
              </TableBody>
            </Table>
          </TableContainer>
        )}
      </Stack>

      {formTarget && (
        <EntityFormDialog
          context={context}
          open
          title={formTarget.mode === 'edit' ? 'Modifica progetto' : 'Nuovo progetto'}
          fields={buildProjectFields(formTarget.mode === 'edit')}
          initialValues={formTarget.mode === 'edit' ? projectToFormValues(formTarget.project) : {}}
          dateRange={{ startField: 'StartDate', endField: 'EndDate' }}
          onCancel={() => setFormTarget(null)}
          onSave={handleSaveProject}
        />
      )}

      <ConfirmDialog
        open={!!deleteTarget}
        title="Elimina progetto"
        message={
          deleteTarget && (
            <>
              Confermi l&apos;eliminazione del progetto <strong>{deleteTarget.ProjectCode} &ndash; {deleteTarget.Title}</strong>?
              <br />
              <br />
              Deliverable, issue, allocazioni e costi collegati a questo progetto <strong>non</strong> verranno eliminati
              automaticamente e resteranno nelle rispettive liste con il codice progetto originario.
            </>
          )
        }
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
