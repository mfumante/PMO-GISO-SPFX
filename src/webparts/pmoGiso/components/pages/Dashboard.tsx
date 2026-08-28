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
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import type { ReferenceLineProps } from 'recharts';
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

// Riga del Gantt: 'offset' e' un Bar trasparente usato solo per spingere
// a destra l'inizio della barra visibile (tecnica standard per un Gantt
// con Recharts, che non ha un tipo di grafico Gantt nativo); 'completed'/
// 'remaining' spezzano la durata del progetto in base a Progress cosi' la
// barra mostra sia la pianificazione (inizio-fine) sia l'avanzamento reale.
interface IGanttRow {
  code: string;
  title: string;
  status: ProjectStatus;
  start: number;
  end: number;
  progress: number;
  offset: number;
  completed: number;
  remaining: number;
}

const currencyFormatter = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const ganttDateFormatter = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: 'short' });
const ganttTooltipDateFormatter = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: 'short', year: 'numeric' });
const DAY_MS = 24 * 60 * 60 * 1000;

// recharts 2.x dichiara ReferenceLine come classe React (React.Component<Props>),
// ma sotto @types/react 17 il tipo risultante non soddisfa l'interfaccia
// JSX.ElementClass attesa dal compilatore (bug di tipizzazione noto di recharts,
// non un problema di runtime: il componente funziona). Il cast e' isolato al solo
// ReferenceLine - gli altri componenti recharts usati in questo file (Bar, XAxis,
// YAxis, Pie, ecc.) non sono affetti e restano tipizzati normalmente.
// Si usa il tipo Props pubblico esportato da recharts (ReferenceLineProps)
// invece di estrarlo dalla classe: React.ComponentProps<typeof ReferenceLine>
// fallisce per lo stesso motivo (l'istanza della classe non e' assegnabile a
// Component<any,any,any> sotto questi @types/react), quindi il tipo va preso
// dall'export pubblico. RefLine resta tipizzato sui props reali, niente 'any'.
const RefLine = ReferenceLine as unknown as React.FC<ReferenceLineProps>;

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

interface GanttTooltipProps {
  active?: boolean;
  payload?: Array<{ payload: IGanttRow }>;
}

// Componente React: il tipo di ritorno di una FunctionComponent e' fissato da
// @types/react a ReactElement | null (non accetta undefined), quindi "null" qui
// e' il modo idiomatico per dire "non renderizzare nulla" quando il tooltip di
// recharts non e' attivo, non un valore di dato modellabile con undefined.
// eslint-disable-next-line @rushstack/no-new-null
function GanttTooltip({ active, payload }: GanttTooltipProps): React.ReactElement | null {
  if (!active || !payload || payload.length === 0) {
    return null;
  }
  const row = payload[0].payload;
  return (
    <Paper elevation={3} sx={{ p: 1.5, maxWidth: 260 }}>
      <Typography variant="subtitle2">
        {row.code} — {row.title}
      </Typography>
      <Typography variant="body2">
        {ganttTooltipDateFormatter.format(new Date(row.start))} → {ganttTooltipDateFormatter.format(new Date(row.end))}
      </Typography>
      <Typography variant="body2">
        Stato: {row.status} · Avanzamento: {Math.round(row.progress)}%
      </Typography>
    </Paper>
  );
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

    const projectsWithDates = projects
      .filter((project) => project.StartDate && project.EndDate)
      .map((project) => ({
        code: project.ProjectCode,
        title: project.Title,
        status: project.Status,
        start: new Date(project.StartDate as string).getTime(),
        end: new Date(project.EndDate as string).getTime(),
        progress: Math.min(100, Math.max(0, project.Progress ?? 0)),
      }))
      .filter((project) => project.end >= project.start)
      .sort((a, b) => a.start - b.start);

    const ganttDomainStart = projectsWithDates.length > 0 ? Math.min(...projectsWithDates.map((p) => p.start)) : 0;
    const ganttDomainEnd = projectsWithDates.length > 0 ? Math.max(...projectsWithDates.map((p) => p.end)) : 0;
    // Se tutti i progetti iniziano e finiscono lo stesso giorno il range
    // sarebbe 0: l'asse X (numerico, in ms) non avrebbe tick leggibili, quindi
    // si forza un range minimo di un giorno.
    const ganttSpan = ganttDomainEnd - ganttDomainStart > 0 ? ganttDomainEnd - ganttDomainStart : DAY_MS;

    const ganttData: IGanttRow[] = projectsWithDates.map((project) => {
      const duration = Math.max(project.end - project.start, DAY_MS);
      const completed = duration * (project.progress / 100);
      return {
        code: project.code,
        title: project.title,
        status: project.status,
        start: project.start,
        end: project.end,
        progress: project.progress,
        offset: project.start - ganttDomainStart,
        completed,
        remaining: duration - completed,
      };
    });

    const todayOffset = Date.now() - ganttDomainStart;
    const projectsMissingDates = projects.length - projectsWithDates.length;

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
      ganttData,
      ganttDomainStart,
      ganttSpan,
      todayOffset,
      projectsMissingDates,
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

            {computed.ganttData.length > 0 && (
              <Paper variant="outlined" sx={{ p: 2 }}>
                <Typography variant="subtitle1" gutterBottom>
                  Pianificazione progetti (Gantt)
                </Typography>
                <Box sx={{ width: '100%', height: Math.max(220, computed.ganttData.length * 44 + 60) }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={computed.ganttData} layout="vertical" margin={{ top: 8, right: 24, left: 8, bottom: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                      <XAxis
                        type="number"
                        domain={[0, computed.ganttSpan]}
                        tickFormatter={(value: number) => ganttDateFormatter.format(new Date(value + computed.ganttDomainStart))}
                      />
                      <YAxis type="category" dataKey="code" width={90} />
                      <Tooltip content={<GanttTooltip />} />
                      <RefLine x={computed.todayOffset} stroke="#A6145C" strokeDasharray="4 4" label={{ value: 'Oggi', position: 'top', fill: '#A6145C', fontSize: 12 }} />
                      <Bar dataKey="offset" stackId="gantt" fill="transparent" isAnimationActive={false} />
                      <Bar dataKey="completed" stackId="gantt" radius={[4, 4, 4, 4]} isAnimationActive={false}>
                        {computed.ganttData.map((entry) => (
                          <Cell key={`completed-${entry.code}`} fill={STATUS_COLORS[entry.status]} />
                        ))}
                      </Bar>
                      <Bar dataKey="remaining" stackId="gantt" radius={[4, 4, 4, 4]} isAnimationActive={false}>
                        {computed.ganttData.map((entry) => (
                          <Cell key={`remaining-${entry.code}`} fill={STATUS_COLORS[entry.status]} fillOpacity={0.28} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </Box>
                <Typography variant="caption" color="text.secondary">
                  Colore pieno = avanzamento completato, colore chiaro = quota residua. Linea tratteggiata = data odierna.
                  {computed.projectsMissingDates > 0 &&
                    ` ${computed.projectsMissingDates} progetti esclusi per data inizio/fine mancante.`}
                </Typography>
              </Paper>
            )}

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
