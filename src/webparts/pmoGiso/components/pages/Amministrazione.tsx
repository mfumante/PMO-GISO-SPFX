import * as React from 'react';
import { useMemo, useState } from 'react';
import {
  Accordion,
  AccordionDetails,
  AccordionSummary,
  Alert,
  Box,
  Button,
  Card,
  CardContent,
  CardHeader,
  Chip,
  CircularProgress,
  Divider,
  LinearProgress,
  Paper,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Typography,
} from '@mui/material';
import BuildOutlinedIcon from '@mui/icons-material/BuildOutlined';
import RefreshOutlinedIcon from '@mui/icons-material/RefreshOutlined';
import DownloadOutlinedIcon from '@mui/icons-material/DownloadOutlined';
import UploadFileOutlinedIcon from '@mui/icons-material/UploadFileOutlined';
import FileDownloadOutlinedIcon from '@mui/icons-material/FileDownloadOutlined';
import PlayArrowOutlinedIcon from '@mui/icons-material/PlayArrowOutlined';
import ExpandMoreOutlinedIcon from '@mui/icons-material/ExpandMoreOutlined';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { PageContainer } from '../layout/PageContainer';
import { FIELD_HELP, HELP_ENTITIES, HelpEntity } from '../help/fieldHelp';
import {
  EnvironmentItemStatus,
  IListStatus,
  SharePointProvisioningService,
} from '../../services/SharePointProvisioningService';
import {
  ExcelService,
  IImportReport,
  IParsedWorkbook,
  IValidationIssue,
  IValidationReport,
  SheetName,
} from '../../services/ExcelService';

interface AmministrazioneProps {
  context: WebPartContext;
}

interface LogEntry {
  id: number;
  time: string;
  level: 'info' | 'success' | 'error';
  message: string;
}

interface EnvironmentRow {
  key: string;
  label: string;
  status: EnvironmentItemStatus;
  message: string;
}

const statusChipProps: Record<EnvironmentItemStatus, { label: string; color: 'success' | 'info' | 'default' | 'error' }> = {
  Existing: { label: 'Existing', color: 'success' },
  Created: { label: 'Created', color: 'info' },
  Missing: { label: 'Missing', color: 'default' },
  Error: { label: 'Error', color: 'error' },
};

const DATA_SHEET_NAMES: SheetName[] = ['Projects', 'Deliverables', 'Issues', 'Resources', 'Allocations', 'Costs'];

interface ISheetSummary {
  sheet: SheetName;
  validCount: number;
  updateCount: number;
  errorCount: number;
  warningCount: number;
}

interface IImportSummary {
  sheet: SheetName;
  created: number;
  updated: number;
  skipped: number;
}

function flattenResults(results: IListStatus[]): EnvironmentRow[] {
  const rows: EnvironmentRow[] = [];
  for (const list of results) {
    rows.push({ key: list.listTitle, label: list.listTitle, status: list.status, message: list.message });
    for (const field of list.fields) {
      rows.push({
        key: `${list.listTitle}.${field.internalName}`,
        label: `${list.listTitle} › ${field.internalName}`,
        status: field.status,
        message: field.message,
      });
    }
  }
  return rows;
}

function buildSheetSummaries(report: IValidationReport): ISheetSummary[] {
  return DATA_SHEET_NAMES.map((sheet) => {
    const result = report[sheet];
    return {
      sheet,
      validCount: result.validRows.length,
      updateCount: result.validRows.filter((row) => row.isUpdate).length,
      errorCount: result.errors.length,
      warningCount: result.warnings.length,
    };
  });
}

function buildAllIssues(report: IValidationReport): IValidationIssue[] {
  const combined: IValidationIssue[] = [];
  DATA_SHEET_NAMES.forEach((sheet) => {
    report[sheet].errors.forEach((issue) => combined.push(issue));
    report[sheet].warnings.forEach((issue) => combined.push(issue));
  });
  return combined;
}

function buildImportSummaries(report: IImportReport): IImportSummary[] {
  return DATA_SHEET_NAMES.map((sheet) => {
    const result = report[sheet];
    return { sheet, created: result.created, updated: result.updated, skipped: result.skipped };
  });
}

function EntityHelpAccordion({ entityKey, label }: { entityKey: HelpEntity; label: string }): React.ReactElement {
  const fields = FIELD_HELP[entityKey];
  const fieldNames = Object.keys(fields);

  return (
    <Accordion disableGutters>
      <AccordionSummary expandIcon={<ExpandMoreOutlinedIcon />}>
        <Typography variant="subtitle2">{label}</Typography>
      </AccordionSummary>
      <AccordionDetails>
        <Stack spacing={1.5} divider={<Divider flexItem />}>
          {fieldNames.map((fieldName) => {
            const help = fields[fieldName];
            return (
              <Box key={fieldName}>
                <Typography variant="body2" sx={{ fontWeight: 700, fontFamily: 'monospace' }}>
                  {fieldName}
                </Typography>
                <Typography variant="body2" color="text.secondary">
                  {help.long}
                </Typography>
                {help.example && (
                  <Typography variant="caption" color="text.secondary" sx={{ fontStyle: 'italic' }}>
                    Esempio: {help.example}
                  </Typography>
                )}
              </Box>
            );
          })}
        </Stack>
      </AccordionDetails>
    </Accordion>
  );
}

export default function Amministrazione({ context }: AmministrazioneProps): React.ReactElement {
  const service = useMemo(() => new SharePointProvisioningService(context), [context]);
  const excelService = useMemo(() => new ExcelService(context), [context]);
  const siteUrl = context.pageContext.web.absoluteUrl;
  const userDisplayName = context.pageContext.user.displayName;

  const [rows, setRows] = useState<EnvironmentRow[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [verifying, setVerifying] = useState(false);
  const [initializing, setInitializing] = useState(false);

  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [parsedWorkbook, setParsedWorkbook] = useState<IParsedWorkbook | null>(null);
  const [validationReport, setValidationReport] = useState<IValidationReport | null>(null);
  const [validating, setValidating] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importProgress, setImportProgress] = useState<{ completed: number; total: number } | null>(null);
  const [importReport, setImportReport] = useState<IImportReport | null>(null);
  const [exporting, setExporting] = useState(false);

  function addLog(level: LogEntry['level'], message: string): void {
    setLogs((prev) => [
      ...prev,
      { id: prev.length, time: new Date().toLocaleTimeString('it-IT'), level, message },
    ]);
  }

  async function handleVerifyEnvironment(): Promise<void> {
    setVerifying(true);
    addLog('info', 'Verifica ambiente SharePoint in corso...');

    try {
      const results = await service.checkEnvironment();
      setRows(flattenResults(results));
      results.forEach((list) => addLog(list.status === 'Error' ? 'error' : 'info', `${list.listTitle}: ${list.message}`));

      const missingCount = results.filter((r) => r.status === 'Missing').length;
      const errorCount = results.filter((r) => r.status === 'Error').length;
      const existingCount = results.length - missingCount - errorCount;
      addLog(
        errorCount > 0 ? 'error' : 'success',
        `Verifica completata: ${existingCount} liste esistenti, ${missingCount} mancanti, ${errorCount} errori.`,
      );
    } catch (error) {
      addLog('error', error instanceof Error ? error.message : 'Errore durante la verifica dell\'ambiente.');
    } finally {
      setVerifying(false);
    }
  }

  async function handleInitializeEnvironment(): Promise<void> {
    setInitializing(true);
    addLog('info', 'Inizializzazione ambiente SharePoint in corso...');

    try {
      const report = await service.initializeEnvironment();
      setRows(flattenResults(report.results));

      report.results.forEach((list) => {
        addLog(list.status === 'Error' ? 'error' : list.status === 'Created' ? 'success' : 'info', `${list.listTitle}: ${list.message}`);
        list.fields
          .filter((field) => field.status === 'Created' || field.status === 'Error')
          .forEach((field) =>
            addLog(
              field.status === 'Error' ? 'error' : 'success',
              `${list.listTitle} › ${field.internalName}: ${field.message}`,
            ),
          );
      });

      addLog(
        report.errorCount > 0 ? 'error' : 'success',
        `Inizializzazione completata: ${report.existingCount} gia' esistenti, ${report.createdCount} create, ${report.errorCount} errori.`,
      );
    } catch (error) {
      addLog('error', error instanceof Error ? error.message : 'Errore durante l\'inizializzazione dell\'ambiente.');
    } finally {
      setInitializing(false);
    }
  }

  function handleDownloadTemplate(): void {
    excelService.generateTemplate();
    addLog('info', 'Template Excel scaricato.');
  }

  async function handleFileSelected(event: React.ChangeEvent<HTMLInputElement>): Promise<void> {
    const fileList = event.target.files;
    const file = fileList && fileList.length > 0 ? fileList[0] : undefined;
    event.target.value = '';

    setSelectedFile(null);
    setParsedWorkbook(null);
    setValidationReport(null);
    setImportReport(null);
    setImportProgress(null);

    if (!file) {
      return;
    }

    setSelectedFile(file);
    setValidating(true);
    addLog('info', `File selezionato: ${file.name}`);

    try {
      const parsed = await excelService.parseWorkbook(file);
      setParsedWorkbook(parsed);
      if (parsed.missingSheets.length > 0) {
        addLog('info', `Fogli non trovati nel file (considerati vuoti): ${parsed.missingSheets.join(', ')}.`);
      }

      const report = await excelService.validateData(parsed);
      setValidationReport(report);

      buildAllIssues(report).forEach((issue) =>
        addLog(issue.level === 'error' ? 'error' : 'info', `[${issue.sheet} riga ${issue.row}] ${issue.message}`),
      );

      addLog(
        report.hasBlockingErrors ? 'error' : 'success',
        report.hasBlockingErrors
          ? 'Validazione completata: sono presenti errori bloccanti, correggi il file prima di importare.'
          : 'Validazione completata: nessun errore bloccante.',
      );
    } catch (error) {
      addLog('error', error instanceof Error ? error.message : 'Errore durante la lettura del file.');
    } finally {
      setValidating(false);
    }
  }

  async function handleImport(): Promise<void> {
    if (!validationReport) {
      return;
    }

    setImporting(true);
    setImportReport(null);
    setImportProgress({ completed: 0, total: 0 });
    addLog('info', 'Import dati in corso...');

    try {
      const report = await excelService.importData(validationReport, {
        onLog: (entry) =>
          addLog(
            entry.level === 'success' ? 'success' : entry.level === 'error' ? 'error' : 'info',
            `[${entry.sheet}${entry.row !== undefined && entry.row !== null ? ` riga ${entry.row}` : ''}] ${entry.message}`,
          ),
        onProgress: (completed, total) => setImportProgress({ completed, total }),
      });
      setImportReport(report);

      let totalErrors = 0;
      DATA_SHEET_NAMES.forEach((sheet) => {
        totalErrors += report[sheet].errors.length;
      });
      addLog(totalErrors > 0 ? 'error' : 'success', `Import completato: ${totalErrors} errori complessivi.`);
    } catch (error) {
      addLog('error', error instanceof Error ? error.message : "Errore durante l'import.");
    } finally {
      setImporting(false);
    }
  }

  async function handleExport(): Promise<void> {
    setExporting(true);
    addLog('info', 'Esportazione dati in corso...');

    try {
      await excelService.exportAllData();
      addLog('success', 'Esportazione completata.');
    } catch (error) {
      addLog('error', error instanceof Error ? error.message : "Errore durante l'esportazione.");
    } finally {
      setExporting(false);
    }
  }

  const sheetSummaries = validationReport ? buildSheetSummaries(validationReport) : [];
  const allIssues = validationReport ? buildAllIssues(validationReport) : [];
  const importSummaries = importReport ? buildImportSummaries(importReport) : [];

  return (
    <PageContainer title="Amministrazione">
      <Stack spacing={3}>
        <Card variant="outlined">
          <CardHeader title="Stato Ambiente SharePoint" subheader="Liste richieste dall'applicazione PMO GISO" />
          <Divider />
          <CardContent>
            <Stack spacing={2}>
              <Stack direction="row" spacing={2} alignItems="center" flexWrap="wrap">
                <Typography variant="body2" color="text.secondary" sx={{ minWidth: 140 }}>
                  Site URL
                </Typography>
                <Typography variant="body2" sx={{ wordBreak: 'break-all' }}>
                  {siteUrl}
                </Typography>
              </Stack>

              <Stack direction="row" spacing={2} alignItems="center" flexWrap="wrap">
                <Typography variant="body2" color="text.secondary" sx={{ minWidth: 140 }}>
                  Utente
                </Typography>
                <Typography variant="body2">{userDisplayName}</Typography>
              </Stack>
            </Stack>

            <Stack direction="row" spacing={2} mt={3} flexWrap="wrap" useFlexGap>
              <Button
                variant="outlined"
                color="primary"
                startIcon={verifying ? <CircularProgress size={16} color="inherit" /> : <RefreshOutlinedIcon />}
                onClick={handleVerifyEnvironment}
                disabled={verifying || initializing}
              >
                Verifica Ambiente
              </Button>
              <Button
                variant="contained"
                color="warning"
                startIcon={initializing ? <CircularProgress size={16} color="inherit" /> : <BuildOutlinedIcon />}
                onClick={handleInitializeEnvironment}
                disabled={verifying || initializing}
              >
                Inizializza Ambiente
              </Button>
            </Stack>

            <Box mt={3}>
              {rows.length === 0 ? (
                <Alert severity="info">
                  Nessuna verifica eseguita. Premi &quot;Verifica Ambiente&quot; per controllare lo stato di liste e
                  colonne.
                </Alert>
              ) : (
                <TableContainer component={Paper} variant="outlined">
                  <Table size="small">
                    <TableHead>
                      <TableRow>
                        <TableCell>Lista</TableCell>
                        <TableCell>Stato</TableCell>
                        <TableCell>Messaggio</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {rows.map((row) => (
                        <TableRow key={row.key}>
                          <TableCell sx={{ fontFamily: 'monospace' }}>{row.label}</TableCell>
                          <TableCell>
                            <Chip size="small" label={statusChipProps[row.status].label} color={statusChipProps[row.status].color} />
                          </TableCell>
                          <TableCell>
                            <Typography variant="body2" color="text.secondary">
                              {row.message}
                            </Typography>
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
              )}
            </Box>
          </CardContent>
        </Card>

        <Card variant="outlined">
          <CardHeader title="Gestione Dati" subheader="Import ed export dei dati PMO GISO da/verso Excel" />
          <Divider />
          <CardContent>
            <Stack spacing={3}>
              <Stack direction="row" spacing={2} flexWrap="wrap" useFlexGap alignItems="center">
                <Button variant="outlined" startIcon={<DownloadOutlinedIcon />} onClick={handleDownloadTemplate}>
                  Scarica template Excel
                </Button>
                <Button variant="outlined" component="label" startIcon={<UploadFileOutlinedIcon />}>
                  Carica file Excel
                  <input type="file" accept=".xlsx" hidden onChange={handleFileSelected} />
                </Button>
                <Button
                  variant="outlined"
                  startIcon={exporting ? <CircularProgress size={16} color="inherit" /> : <FileDownloadOutlinedIcon />}
                  onClick={handleExport}
                  disabled={exporting}
                >
                  Esporta tutti i dati
                </Button>
              </Stack>

              {selectedFile && (
                <Typography variant="body2" color="text.secondary">
                  File selezionato: <strong>{selectedFile.name}</strong>
                </Typography>
              )}

              {validating && <LinearProgress />}

              {parsedWorkbook && parsedWorkbook.missingSheets.length > 0 && (
                <Alert severity="warning">
                  Fogli non trovati nel file (considerati vuoti): {parsedWorkbook.missingSheets.join(', ')}.
                </Alert>
              )}

              {validationReport && (
                <>
                  <Box>
                    <Typography variant="subtitle2" gutterBottom>
                      Anteprima record per foglio
                    </Typography>
                    <TableContainer component={Paper} variant="outlined">
                      <Table size="small">
                        <TableHead>
                          <TableRow>
                            <TableCell>Foglio</TableCell>
                            <TableCell align="right">Righe valide</TableCell>
                            <TableCell align="right">Di cui aggiornamenti</TableCell>
                            <TableCell align="right">Errori</TableCell>
                            <TableCell align="right">Warning</TableCell>
                          </TableRow>
                        </TableHead>
                        <TableBody>
                          {sheetSummaries.map((summary) => (
                            <TableRow key={summary.sheet}>
                              <TableCell>{summary.sheet}</TableCell>
                              <TableCell align="right">{summary.validCount}</TableCell>
                              <TableCell align="right">{summary.updateCount}</TableCell>
                              <TableCell align="right">{summary.errorCount}</TableCell>
                              <TableCell align="right">{summary.warningCount}</TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </TableContainer>
                  </Box>

                  {allIssues.length > 0 && (
                    <Box>
                      <Typography variant="subtitle2" gutterBottom>
                        Validazione
                      </Typography>
                      <Paper variant="outlined" sx={{ p: 2, maxHeight: 280, overflowY: 'auto' }}>
                        <Stack spacing={0.5}>
                          {allIssues.map((issue, index) => (
                            <Typography
                              key={index}
                              variant="body2"
                              sx={{ color: issue.level === 'error' ? 'error.main' : 'warning.main' }}
                            >
                              [{issue.sheet} riga {issue.row}] {issue.message}
                            </Typography>
                          ))}
                        </Stack>
                      </Paper>
                    </Box>
                  )}

                  <Box>
                    <Button
                      variant="contained"
                      color="warning"
                      startIcon={importing ? <CircularProgress size={16} color="inherit" /> : <PlayArrowOutlinedIcon />}
                      onClick={handleImport}
                      disabled={importing || validationReport.hasBlockingErrors}
                    >
                      Esegui Import
                    </Button>
                  </Box>

                  {importProgress && (
                    <Box>
                      <LinearProgress
                        variant="determinate"
                        value={importProgress.total > 0 ? (importProgress.completed / importProgress.total) * 100 : 0}
                      />
                      <Typography variant="caption" color="text.secondary">
                        {importProgress.completed} / {importProgress.total}
                      </Typography>
                    </Box>
                  )}

                  {importReport && (
                    <Box>
                      <Typography variant="subtitle2" gutterBottom>
                        Report import
                      </Typography>
                      <TableContainer component={Paper} variant="outlined">
                        <Table size="small">
                          <TableHead>
                            <TableRow>
                              <TableCell>Lista</TableCell>
                              <TableCell align="right">Creati</TableCell>
                              <TableCell align="right">Aggiornati</TableCell>
                              <TableCell align="right">Saltati</TableCell>
                            </TableRow>
                          </TableHead>
                          <TableBody>
                            {importSummaries.map((summary) => (
                              <TableRow key={summary.sheet}>
                                <TableCell>{summary.sheet}</TableCell>
                                <TableCell align="right">{summary.created}</TableCell>
                                <TableCell align="right">{summary.updated}</TableCell>
                                <TableCell align="right">{summary.skipped}</TableCell>
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </TableContainer>
                    </Box>
                  )}
                </>
              )}
            </Stack>
          </CardContent>
        </Card>

        <Card variant="outlined">
          <CardHeader title="Guida alla compilazione" subheader="Significato dei campi per ciascuna entita'" />
          <Divider />
          <CardContent>
            <Stack spacing={1}>
              {HELP_ENTITIES.map((entity) => (
                <EntityHelpAccordion key={entity.key} entityKey={entity.key} label={entity.label} />
              ))}
            </Stack>
          </CardContent>
        </Card>

        <Card variant="outlined">
          <CardHeader title="Log" subheader="Dettaglio delle operazioni eseguite" />
          <Divider />
          <CardContent>
            <Paper
              variant="outlined"
              sx={{
                p: 2,
                maxHeight: 320,
                overflowY: 'auto',
                bgcolor: 'grey.50',
                fontFamily: 'monospace',
              }}
            >
              {logs.length === 0 ? (
                <Typography variant="body2" color="text.secondary">
                  Nessuna operazione eseguita.
                </Typography>
              ) : (
                <Stack spacing={0.5}>
                  {logs.map((log) => (
                    <Typography
                      key={log.id}
                      variant="body2"
                      component="div"
                      sx={{
                        fontFamily: 'monospace',
                        color:
                          log.level === 'error'
                            ? 'error.main'
                            : log.level === 'success'
                              ? 'success.main'
                              : 'text.secondary',
                      }}
                    >
                      [{log.time}] {log.message}
                    </Typography>
                  ))}
                </Stack>
              )}
            </Paper>
          </CardContent>
        </Card>
      </Stack>
    </PageContainer>
  );
}
