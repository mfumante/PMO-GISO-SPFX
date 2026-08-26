import * as React from 'react';
import { useMemo, useState } from 'react';
import {
  Alert,
  Button,
  IconButton,
  Skeleton,
  Snackbar,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Tooltip,
  Typography,
} from '@mui/material';
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
import { useLayoutPreferences } from '../layout/LayoutPreferencesContext';
import { getTableSx, TABLE_MAX_HEIGHT } from '../common/tableStyles';
import { buildResourceFields } from '../forms/entityFormSchemas';
import { formValuesToNewResource, resolvePersonFieldId, resourceToFormValues } from '../forms/entityFormValues';
import { PeopleService } from '../../services/PeopleService';
import { IAllocation, IProject, IResource, SharePointDataService } from '../../services/SharePointDataService';

interface RisorseProps {
  context: WebPartContext;
}

interface IResourceMatrixData {
  resources: IResource[];
  allocations: IAllocation[];
  projects: IProject[];
}

type ResourceFormTarget = { mode: 'new' } | { mode: 'edit'; resource: IResource };

// Rampa sequenziale (blu, chiaro->scuro) validata in references/palette.md della skill
// dataviz: qui codifica la magnitudine dell'allocazione (0-100%). Oltre il 100% si passa
// al rosso 'critical' della status palette della stessa skill, riservato a segnalare la
// sovra-allocazione: e' uno stato, non un altro passo della sequenza.
const HEAT_STEPS = ['#cde2fb', '#9ec5f4', '#6da7ec', '#3987e5', '#256abf'];
const CRITICAL_COLOR = '#d03b3b';

function getHeatColor(percent: number): string {
  if (percent <= 0) {
    return 'transparent';
  }
  if (percent > 100) {
    return 'rgba(208, 59, 59, 0.16)';
  }
  const index = Math.min(HEAT_STEPS.length - 1, Math.floor((percent / 100) * HEAT_STEPS.length));
  return HEAT_STEPS[index];
}

export default function Risorse({ context }: RisorseProps): React.ReactElement {
  const service = useMemo(() => new SharePointDataService(context), [context]);
  const peopleService = useMemo(() => new PeopleService(context), [context]);
  const { feedback, showSuccess, showError, closeFeedback } = useFeedback();
  const { density } = useLayoutPreferences();

  const [refreshKey, setRefreshKey] = useState(0);
  const [formTarget, setFormTarget] = useState<ResourceFormTarget | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<IResource | null>(null);
  const [deleting, setDeleting] = useState(false);

  const result = useAsyncData<IResourceMatrixData>(async () => {
    const [resources, allocations, projects] = await Promise.all([
      service.getResources(),
      service.getAllocations(),
      service.getProjects(),
    ]);
    return { resources, allocations, projects };
  }, [service, refreshKey]);

  const matrix = useMemo(() => {
    if (result.status !== 'ready') {
      return undefined;
    }
    const { resources, allocations, projects } = result.data;

    const projectTitleByCode = new Map(
      projects.map((project): [string, string] => [project.ProjectCode, project.Title]),
    );

    // Niente Array.from/Set iterato: la lib TS di questo progetto esclude
    // es2015.core/es2015.iterable, quindi la deduplica resta su array/indexOf (ES5).
    const projectCodes: string[] = [];
    allocations.forEach((allocation) => {
      if (projectCodes.indexOf(allocation.ProjectCode) === -1) {
        projectCodes.push(allocation.ProjectCode);
      }
    });
    projectCodes.sort();

    const rows = resources.map((resource) => {
      const percentByProject = new Map<string, number>();
      let total = 0;
      allocations
        .filter((allocation) => allocation.ResourceCode === resource.ResourceCode)
        .forEach((allocation) => {
          const percent = allocation.AllocationPercent ?? 0;
          const current = percentByProject.get(allocation.ProjectCode) ?? 0;
          percentByProject.set(allocation.ProjectCode, current + percent);
          total += percent;
        });
      return { resource, percentByProject, total };
    });

    return { projectCodes, projectTitleByCode, rows };
  }, [result]);

  function countAllocationsForResource(resourceCode: string): number {
    if (result.status !== 'ready') {
      return 0;
    }
    return result.data.allocations.filter((allocation) => allocation.ResourceCode === resourceCode).length;
  }

  async function handleSaveResource(values: FormValues): Promise<void> {
    const payload = formValuesToNewResource(values);
    payload.PersonUserId = await resolvePersonFieldId(peopleService, values, 'PersonUser');
    const target = formTarget;
    const writeResult =
      target?.mode === 'edit' ? await service.updateResource(target.resource.Id, payload) : await service.createResource(payload);

    if (!writeResult.success) {
      throw new Error(writeResult.message);
    }

    setRefreshKey((key) => key + 1);
    showSuccess(target?.mode === 'edit' ? 'Risorsa aggiornata.' : 'Risorsa creata.');
    setFormTarget(null);
  }

  async function handleConfirmDelete(): Promise<void> {
    if (!deleteTarget) {
      return;
    }
    setDeleting(true);
    try {
      const writeResult = await service.deleteResource(deleteTarget.Id);
      if (!writeResult.success) {
        showError(writeResult.message);
        return;
      }
      setRefreshKey((key) => key + 1);
      showSuccess('Risorsa eliminata.');
      setDeleteTarget(null);
    } catch (error) {
      showError(error instanceof Error ? error.message : 'Errore durante l\'eliminazione.');
    } finally {
      setDeleting(false);
    }
  }

  const deleteAllocationCount = deleteTarget ? countAllocationsForResource(deleteTarget.ResourceCode) : 0;

  return (
    <PageContainer
      title="Risorse"
      actions={
        <Button variant="contained" startIcon={<AddOutlinedIcon />} onClick={() => setFormTarget({ mode: 'new' })}>
          Nuova risorsa
        </Button>
      }
    >
      <Stack spacing={2}>

        {result.status === 'loading' && (
          <Stack spacing={1}>
            {[1, 2, 3, 4].map((key) => (
              <Skeleton key={key} variant="rectangular" height={44} />
            ))}
          </Stack>
        )}

        {result.status === 'missing' && (
          <Alert severity="warning">
            La lista <strong>{result.listTitle}</strong> non esiste ancora. Vai in <strong>Amministrazione</strong> per
            inizializzarla.
          </Alert>
        )}

        {result.status === 'error' && <Alert severity="error">{result.message}</Alert>}

        {result.status === 'ready' && result.data.resources.length === 0 && (
          <Alert severity="info">
            Non ci sono ancora risorse in <strong>PMO_Resources</strong>. Vai in <strong>Amministrazione</strong> per
            importare i dati, oppure crea la prima risorsa con il pulsante &quot;Nuova risorsa&quot;.
          </Alert>
        )}

        {result.status === 'ready' && result.data.resources.length > 0 && (
          <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
            <Table size="small" sx={getTableSx(density, 760)}>
              <TableHead>
                <TableRow>
                  <TableCell sx={{ width: '14%' }}>ResourceCode</TableCell>
                  <TableCell sx={{ width: '28%' }}>Nome</TableCell>
                  <TableCell sx={{ width: '16%' }}>Ruolo</TableCell>
                  <TableCell sx={{ width: '14%' }}>Unita&apos;</TableCell>
                  <TableCell align="right" sx={{ width: '10%' }}>Capacita&apos;</TableCell>
                  <TableCell sx={{ width: '9%' }}>Attiva</TableCell>
                  <TableCell align="right" sx={{ width: '9%' }}>Azioni</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {result.data.resources.map((resource) => (
                  <TableRow key={resource.Id}>
                    <TableCell sx={{ fontFamily: 'monospace' }}>{resource.ResourceCode}</TableCell>
                    <TableCell>
                      <PersonDisplay person={resource.PersonUser} fallbackText={resource.Title} />
                    </TableCell>
                    <TableCell>
                      <ExpandableCellText text={resource.Role} />
                    </TableCell>
                    <TableCell>
                      <ExpandableCellText text={resource.Unit} />
                    </TableCell>
                    <TableCell align="right">{resource.Capacity ?? '-'}</TableCell>
                    <TableCell>{resource.Active ? 'Si' : 'No'}</TableCell>
                    <TableCell align="right">
                      <Tooltip title="Modifica">
                        <IconButton size="small" onClick={() => setFormTarget({ mode: 'edit', resource })}>
                          <EditOutlinedIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="Elimina">
                        <IconButton size="small" onClick={() => setDeleteTarget(resource)}>
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

        {result.status === 'ready' && result.data.resources.length > 0 && matrix && (
          <>
            {matrix.projectCodes.length === 0 ? (
              <Alert severity="info">Nessuna allocazione registrata in PMO_Allocations.</Alert>
            ) : (
              <TableContainer sx={{ maxHeight: TABLE_MAX_HEIGHT }}>
                <Table size="small" sx={getTableSx(density, 220 + matrix.projectCodes.length * 90 + 90)}>
                  <TableHead>
                    <TableRow>
                      <TableCell sx={{ width: 220 }}>Risorsa</TableCell>
                      {matrix.projectCodes.map((code) => (
                        <TableCell key={code} align="center" title={matrix.projectTitleByCode.get(code) ?? code}>
                          {code}
                        </TableCell>
                      ))}
                      <TableCell align="center" sx={{ width: 90 }}>Totale</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {matrix.rows.map(({ resource, percentByProject, total }) => (
                      <TableRow key={resource.Id}>
                        <TableCell title={`${resource.Title} (${resource.ResourceCode})`}>
                          {resource.Title} <Typography component="span" variant="caption" color="text.secondary">({resource.ResourceCode})</Typography>
                        </TableCell>
                        {matrix.projectCodes.map((code) => {
                          const percent = percentByProject.get(code) ?? 0;
                          return (
                            <TableCell
                              key={code}
                              align="center"
                              sx={{
                                bgcolor: getHeatColor(percent),
                                color: percent > 100 ? CRITICAL_COLOR : undefined,
                                fontWeight: percent > 100 ? 700 : 400,
                              }}
                            >
                              {percent > 0 ? `${percent}%` : '-'}
                            </TableCell>
                          );
                        })}
                        <TableCell
                          align="center"
                          sx={{
                            fontWeight: 700,
                            color: total > 100 ? CRITICAL_COLOR : undefined,
                            bgcolor: total > 100 ? 'rgba(208, 59, 59, 0.12)' : undefined,
                          }}
                        >
                          {total}%
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            )}
          </>
        )}
      </Stack>

      {formTarget && (
        <EntityFormDialog
          context={context}
          open
          title={formTarget.mode === 'edit' ? 'Modifica risorsa' : 'Nuova risorsa'}
          fields={buildResourceFields(formTarget.mode === 'edit')}
          initialValues={formTarget.mode === 'edit' ? resourceToFormValues(formTarget.resource) : {}}
          onCancel={() => setFormTarget(null)}
          onSave={handleSaveResource}
        />
      )}

      <ConfirmDialog
        open={!!deleteTarget}
        title="Elimina risorsa"
        message={
          deleteTarget && (
            <>
              Confermi l&apos;eliminazione della risorsa <strong>{deleteTarget.ResourceCode} &ndash; {deleteTarget.Title}</strong>?
              {deleteAllocationCount > 0 && (
                <>
                  <br />
                  <br />
                  Attenzione: questa risorsa ha <strong>{deleteAllocationCount}</strong> allocazion{deleteAllocationCount === 1 ? 'e' : 'i'}{' '}
                  collegat{deleteAllocationCount === 1 ? 'a' : 'e'} che non verr{deleteAllocationCount === 1 ? 'a' : 'anno'} eliminat
                  {deleteAllocationCount === 1 ? 'a' : 'e'} automaticamente.
                </>
              )}
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
