import * as React from 'react';
import { useMemo } from 'react';
import { Alert, Box, Paper, Skeleton, Stack, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Typography } from '@mui/material';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { PageContainer } from '../layout/PageContainer';
import { KpiCard } from '../common/KpiCard';
import { ExpandableCellText } from '../common/ExpandableCellText';
import { getTableSx } from '../common/tableStyles';
import { useAsyncData } from '../../hooks/useAsyncData';
import { useLayoutPreferences } from '../layout/LayoutPreferencesContext';
import type { WidthTier } from '../layout/LayoutPreferencesContext';
import { IAllocation, IIssue, IProject, IssueSeverity, ProjectStatus, SharePointDataService } from '../../services/SharePointDataService';

interface DashboardProps {
  context: WebPartContext;
}

interface IDashboardData {
  projects: IProject[];
  issues: IIssue[];
  allocations: IAllocation[];
}

const currencyFormatter = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });

const ACTIVE_STATUSES: ProjectStatus[] = ['Not Started', 'In Progress', 'On Hold'];

// Palette categorica validata (skill dataviz, references/palette.md, slot 1-5 in
// ordine fisso) per la torta di distribuzione Status: identita' nominale, non un
// ranking, quindi non va confusa con la palette di brand (#0B2545/#A6145C/#E4572E)
// usata nel resto della UI.
const STATUS_COLORS: Record<ProjectStatus, string> = {
  'Not Started': '#2a78d6',
  'In Progress': '#1baf7a',
  Completed: '#eda100',
  'On Hold': '#008300',
  Cancelled: '#4a3aa7',
};

// Rampa ordinale (una sola tinta, chiaro->scuro) per Severity: la gravita' e' un
// ordine, non un'identita' categorica, quindi un'unica tinta crescente e' la codifica
// corretta (vedi color-formula.md della skill dataviz).
const SEVERITY_COLORS: Record<IssueSeverity, string> = {
  Low: '#86b6ef',
  Medium: '#5598e7',
  High: '#256abf',
  Critical: '#184f95',
};

const SEVERITY_ORDER: IssueSeverity[] = ['Low', 'Medium', 'High', 'Critical'];
const STATUS_ORDER: ProjectStatus[] = ['Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled'];

// Numero di colonne della griglia KPI/grafici in base alla fascia di
// larghezza del contenitore (vedi INTERVENTO 4): 5 KPI su una sola riga in
// fascia ampia, 4 in fascia media, 2 impilate a coppie in fascia stretta.
function kpiColumns(tier: WidthTier): number {
  if (tier === 'narrow') return 2;
  if (tier === 'medium') return 4;
  return 5;
}

function chartColumns(tier: WidthTier): number {
  return tier === 'narrow' ? 1 : 2;
}

export default function Dashboard({ context }: DashboardProps): React.ReactElement {
  const service = useMemo(() => new SharePointDataService(context), [context]);
  const { widthTier, density } = useLayoutPreferences();

  const result = useAsyncData<IDashboardData>(async () => {
    const [projects, issues, allocations] = await Promise.all([
      service.getProjects(),
      service.getAllIssues(),
      service.getAllocations(),
    ]);
    return { projects, issues, allocations };
  }, [service]);

  const computed = useMemo(() => {
    if (result.status !== 'ready') {
      return undefined;
    }
    const { projects, issues, allocations } = result.data;

    const activeProjectsCount = projects.filter((project) => ACTIVE_STATUSES.indexOf(project.Status) !== -1).length;
    const budgetTotal = projects.reduce((sum, project) => sum + (project.BudgetTotal ?? 0), 0);
    const budgetConsumed = projects.reduce((sum, project) => sum + (project.BudgetConsumed ?? 0), 0);
    const openIssuesCount = issues.filter((issue) => issue.Status !== 'Closed').length;
    const involvedResourcesCount = new Set(allocations.map((allocation) => allocation.ResourceCode)).size;

    const statusData = STATUS_ORDER.map((status) => ({
      status,
      count: projects.filter((project) => project.Status === status).length,
    })).filter((entry) => entry.count > 0);

    const progressData = projects
      .slice()
      .sort((a, b) => a.ProjectCode.localeCompare(b.ProjectCode))
      .map((project) => ({ code: project.ProjectCode, progress: Math.round(project.Progress ?? 0) }));

    const severityData = SEVERITY_ORDER.map((severity) => ({
      severity,
      count: issues.filter((issue) => issue.Severity === severity).length,
    }));

    const topProjectsByBudget = projects
      .slice()
      .sort((a, b) => (b.BudgetTotal ?? 0) - (a.BudgetTotal ?? 0))
      .slice(0, 5);

    return {
      activeProjectsCount,
      budgetTotal,
      budgetConsumed,
      openIssuesCount,
      involvedResourcesCount,
      statusData,
      progressData,
      severityData,
      topProjectsByBudget,
    };
  }, [result]);

  return (
    <PageContainer title="Dashboard">
      <Stack spacing={3}>
        {result.status === 'loading' && (
          <Stack spacing={1}>
            <Skeleton variant="rectangular" height={100} />
            <Skeleton variant="rectangular" height={300} />
          </Stack>
        )}

        {result.status === 'missing' && (
          <Alert severity="warning">
            La lista <strong>{result.listTitle}</strong> non esiste ancora. Vai in <strong>Amministrazione</strong> per
            inizializzarla.
          </Alert>
        )}

        {result.status === 'error' && <Alert severity="error">{result.message}</Alert>}

        {result.status === 'ready' && result.data.projects.length === 0 && (
          <Alert severity="info">
            Non ci sono ancora progetti in <strong>PMO_Projects</strong>. Vai in <strong>Amministrazione</strong> per
            importare i dati.
          </Alert>
        )}

        {result.status === 'ready' && computed && result.data.projects.length > 0 && (
          <>
            <Box
              sx={{
                display: 'grid',
                gridTemplateColumns: `repeat(${kpiColumns(widthTier)}, 1fr)`,
                gap: 2,
              }}
            >
              <KpiCard label="Progetti attivi" value={computed.activeProjectsCount} />
              <KpiCard label="Budget totale" value={currencyFormatter.format(computed.budgetTotal)} />
              <KpiCard label="Budget consumato" value={currencyFormatter.format(computed.budgetConsumed)} />
              <KpiCard label="Issue aperte" value={computed.openIssuesCount} />
              <KpiCard label="Risorse coinvolte" value={computed.involvedResourcesCount} />
            </Box>

            <Box
              sx={{
                display: 'grid',
                gridTemplateColumns: `repeat(${chartColumns(widthTier)}, 1fr)`,
                gap: 2,
              }}
            >
              <Paper variant="outlined" sx={{ p: 2 }}>
                <Typography variant="subtitle1" gutterBottom>
                  Distribuzione progetti per Status
                </Typography>
                <Box sx={{ width: '100%', aspectRatio: '16 / 10' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={computed.statusData} dataKey="count" nameKey="status" cx="50%" cy="50%" outerRadius={90}>
                        {computed.statusData.map((entry) => (
                          <Cell key={entry.status} fill={STATUS_COLORS[entry.status]} />
                        ))}
                      </Pie>
                      <Tooltip />
                      <Legend />
                    </PieChart>
                  </ResponsiveContainer>
                </Box>
              </Paper>

              <Paper variant="outlined" sx={{ p: 2 }}>
                <Typography variant="subtitle1" gutterBottom>
                  Issue per Severity
                </Typography>
                <Box sx={{ width: '100%', aspectRatio: '16 / 10' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={computed.severityData}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="severity" />
                      <YAxis allowDecimals={false} />
                      <Tooltip />
                      <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                        {computed.severityData.map((entry) => (
                          <Cell key={entry.severity} fill={SEVERITY_COLORS[entry.severity]} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </Box>
              </Paper>
            </Box>

            <Paper variant="outlined" sx={{ p: 2 }}>
              <Typography variant="subtitle1" gutterBottom>
                Avanzamento per progetto
              </Typography>
              <Box sx={{ width: '100%', aspectRatio: widthTier === 'narrow' ? '4 / 3' : '21 / 6' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={computed.progressData}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} />
                    <XAxis dataKey="code" interval={0} angle={-35} textAnchor="end" height={70} />
                    <YAxis domain={[0, 100]} unit="%" />
                    <Tooltip />
                    <Bar dataKey="progress" name="Progress" fill="#2a78d6" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Box>
            </Paper>

            <Paper variant="outlined" sx={{ p: 2 }}>
              <Typography variant="subtitle1" gutterBottom>
                Top 5 progetti per budget
              </Typography>
              <TableContainer>
                <Table size="small" sx={getTableSx(density, 420)}>
                  <TableHead>
                    <TableRow>
                      <TableCell sx={{ width: '20%' }}>ProjectCode</TableCell>
                      <TableCell sx={{ width: '55%' }}>Title</TableCell>
                      <TableCell align="right" sx={{ width: '25%' }}>BudgetTotal</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {computed.topProjectsByBudget.map((project) => (
                      <TableRow key={project.Id}>
                        <TableCell sx={{ fontFamily: 'monospace' }}>{project.ProjectCode}</TableCell>
                        <TableCell>
                          <ExpandableCellText text={project.Title} />
                        </TableCell>
                        <TableCell align="right">{currencyFormatter.format(project.BudgetTotal ?? 0)}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            </Paper>
          </>
        )}
      </Stack>
    </PageContainer>
  );
}
