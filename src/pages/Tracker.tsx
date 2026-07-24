import React, { useState, useRef, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { format } from 'date-fns';
import { Download, UploadCloud, Search, Copy, Check, FileSpreadsheet, RefreshCw, XCircle, AlertCircle, AlertTriangle, CheckCircle2, Play, ArrowLeft, Mail, Send, Printer, ChevronDown, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Progress } from '@/components/ui/progress';
import { Badge } from '@/components/ui/badge';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';
import { Textarea } from '@/components/ui/textarea';
import { Toaster } from '@/components/ui/sonner';
import { toast } from 'sonner';

// --- Configuration -----------------------------------------------------
const HARDCODED_KEY = 'ed2030c6-2c5e-4a4c-90a7-597656156886';
const API_BASE = 'https://api.company-information.service.gov.uk';
const EMAIL_WEBHOOK_URL = 'https://hook.eu2.make.com/lzvt75ds685qmito8cx5ilrvywndj2dr';
const TEST_EMAIL = 'admin@alaccountingsolutions.com';

// Accounts take longer to prepare than confirmation statements, so this tracker
// uses a wider amber window (90 days) than the Confirmation Statement Tracker (30 days).
// Pending confirmation from the AL Accounting admin team — change this one constant if needed.
const AMBER_THRESHOLD_DAYS = 90;

function padNum(n: string | number): string {
  return String(n).trim().replace(/\s/g, '').padStart(8, '0');
}

// MRet is recorded inconsistently across spreadsheet versions (Y, y, Yes, YES) —
// treat any of those as a retainer flag rather than matching one exact string.
function isRetainer(mret?: string): boolean {
  const v = (mret || '').trim().toLowerCase();
  return v === 'y' || v === 'yes';
}

async function fetchCompany(num: string | number): Promise<{ data?: Record<string, unknown>; error?: string; padded: string }> {
  const padded = padNum(num);
  const url = `${API_BASE}/company/${padded}`;
  try {
    const res = await fetch(url, {
      headers: { 'Authorization': 'Basic ' + btoa(`${HARDCODED_KEY}:`) }
    });
    if (!res.ok) {
      if (res.status === 404) return { error: 'Company not found', padded };
      if (res.status === 401) return { error: 'Invalid API key', padded };
      if (res.status === 403) return { error: 'Forbidden — domain not whitelisted for this API key', padded };
      if (res.status === 429) return { error: 'Rate limit — try again shortly', padded };
      return { error: `HTTP ${res.status}`, padded };
    }
    const data = await res.json();
    return { data, padded };
  } catch {
    return { error: 'Network error', padded };
  }
}

type ResultStatus = 'overdue' | 'soon' | 'clear' | 'error';
interface ClientResult {
  name: string;
  number: string;
  dueDate: string | null;
  dueObj: Date | null;
  madeUpTo: string | null;
  madeUpToObj: Date | null;
  diffDays: number | null;
  status: ResultStatus;
  error?: string;
  contactName?: string;
  email?: string;
  mret?: string;
  periodStart?: string;
  periodEnd?: string;
}
interface MismatchResult {
  name: string;
  number: string;
  ourStatus: string;
  chStatus: string;
}

const delay = (ms: number) => new Promise(r => setTimeout(r, ms));

// Delay between each webhook call while the email send queue drains, so Make.com
// never receives a burst of rapid-fire requests when staff queue up several rows.
const EMAIL_QUEUE_DELAY_MS = 3000;
const EMAIL_QUEUE_CLEAR_DELAY_MS = 4000;

type EmailQueueStatus = 'queued' | 'sending' | 'sent' | 'failed';
interface EmailQueueItem {
  id: string;
  number: string;
  name: string;
  to: string;
  subject: string;
  bodyHtml: string;
  status: EmailQueueStatus;
  error?: string;
}

function formatPeriodDate(d?: string): string {
  if (!d) return '';
  const parsed = new Date(d);
  return isNaN(parsed.getTime()) ? d : format(parsed, 'dd MMM yyyy');
}

// Two templates selected automatically by MRet status — retainer clients have the
// deposit request line removed since they're already on a monthly retainer.
// Plain text with emoji/bullet markers (not HTML tags) so the draft stays a normal,
// reliably editable Textarea — converted to HTML only at send time.
function buildEmailBody(r: ClientResult): string {
  const firstName = r.contactName ? r.contactName.split(' ')[0] : r.name.split(' ')[0];
  const periodStart = formatPeriodDate(r.periodStart);
  const periodEnd = formatPeriodDate(r.periodEnd);
  const dueDate = formatPeriodDate(r.dueDate || undefined);

  const depositLine = isRetainer(r.mret) ? '' : `
💷 Please pay an advance deposit of £120 (Bank: AL Accounting Solutions Ltd, Sort code: 60-02-38, Account: 68132328, Reference: your company name)`;

  return `Dear ${firstName},

🏢 Company Name: ${r.name}
📅 Accounting period: ${periodStart} to ${periodEnd}
⏰ Due date: ${dueDate}

Your company's accounts and corporation tax return are due. We require the following documents:

✅ Company bank transactions for the period (please download transactions to Excel/CSV and PDF and send to this email)
✅ Sales or income schedule
✅ Purchases or expenses schedule for the period
✅ Payroll information for the period
✅ Any other documents that will help us complete the accounts${depositLine}

We look forward to hearing from you soon!

Kind Regards,
Florence D Mandevane
Accounts Assistant
AL Accounting Solutions
admin@alaccountingsolutions.com
www.alaccountingsolutions.com`;
}

export default function Tracker() {
  const [allResults, setAllResults] = useState<ClientResult[]>([]);
  const [currentFilter, setCurrentFilter] = useState<'all' | 'overdue' | 'soon' | 'clear' | 'error'>('all');
  const [search, setSearch] = useState('');
  const cancelledRef = useRef<boolean>(false);
  const [rawData, setRawData] = useState<string[][]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [phase, setPhase] = useState<'setup' | 'mapped' | 'running' | 'done'>('setup');

  const [progressCurrent, setProgressCurrent] = useState(0);
  const [progressTotal, setProgressTotal] = useState(0);
  const [progressLabel, setProgressLabel] = useState('');

  const [successMsg, setSuccessMsg] = useState('');
  const [errorMsg, setErrorMsg] = useState('');

  const [copied, setCopied] = useState(false);

  // Defaults to the end of the current month — this is run at the start of each
  // month to catch every client whose accounts are outstanding up to that month's end.
  const now = new Date();
  const defaultMadeUpTo = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const [madeUpToDate, setMadeUpToDate] = useState<string>(defaultMadeUpTo.toISOString().slice(0, 10));

  const [filename, setFilename] = useState('');
  const [nameColIdx, setNameColIdx] = useState<string>('');
  const [numberColIdx, setNumberColIdx] = useState<string>('');
  const [statusColIdx, setStatusColIdx] = useState<number>(-1);
  const [contactNameColIdx, setContactNameColIdx] = useState<number>(-1);
  const [emailColIdx, setEmailColIdx] = useState<number>(-1);
  const [mretColIdx, setMretColIdx] = useState<number>(-1);

  const [mismatches, setMismatches] = useState<MismatchResult[]>([]);
  const [mismatchOpen, setMismatchOpen] = useState(false);

  // Email send queue state
  const [emailQueue, setEmailQueue] = useState<EmailQueueItem[]>([]);
  const emailQueueRef = useRef<EmailQueueItem[]>([]);
  const emailQueueProcessingRef = useRef(false);

  // Email modal state
  const [emailModalOpen, setEmailModalOpen] = useState(false);
  const [modalRow, setModalRow] = useState<ClientResult | null>(null);
  const [draftTo, setDraftTo] = useState('');
  const [draftSubject, setDraftSubject] = useState('');
  const [draftBody, setDraftBody] = useState('');

  const updateEmailQueue = (updater: (prev: EmailQueueItem[]) => EmailQueueItem[]) => {
    const next = updater(emailQueueRef.current);
    emailQueueRef.current = next;
    setEmailQueue(next);
  };

  const processEmailQueue = async () => {
    if (emailQueueProcessingRef.current) return;
    emailQueueProcessingRef.current = true;
    while (true) {
      const next = emailQueueRef.current.find(i => i.status === 'queued');
      if (!next) break;
      updateEmailQueue(prev => prev.map(i => i.id === next.id ? { ...i, status: 'sending' } : i));

      let success = false;
      let errMsg = '';
      if (EMAIL_WEBHOOK_URL) {
        try {
          await fetch(EMAIL_WEBHOOK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ to: next.to, subject: next.subject, body: next.bodyHtml }),
          });
          success = true;
        } catch {
          errMsg = 'Webhook request failed — check the Make.com URL / connection.';
        }
      } else {
        errMsg = 'Make.com webhook not configured yet — set EMAIL_WEBHOOK_URL in Tracker.tsx.';
      }

      updateEmailQueue(prev => prev.map(i => i.id === next.id
        ? { ...i, status: success ? 'sent' : 'failed', error: success ? undefined : errMsg }
        : i));

      if (success) {
        toast.success(`Email sent — ${next.name}`, { description: `To: ${next.to}` });
      } else {
        toast.error(`Email failed — ${next.name}`, { description: errMsg });
      }

      if (emailQueueRef.current.some(i => i.status === 'queued')) {
        await delay(EMAIL_QUEUE_DELAY_MS);
      }
    }
    emailQueueProcessingRef.current = false;
  };

  const enqueueEmail = (item: Omit<EmailQueueItem, 'status' | 'error'>) => {
    updateEmailQueue(prev => [...prev.filter(i => i.status !== 'sent'), { ...item, status: 'queued' }]);
    toast.success(`Email queued — ${item.name}`);
    void processEmailQueue();
  };

  const retryEmail = (id: string) => {
    updateEmailQueue(prev => prev.map(i => i.id === id ? { ...i, status: 'queued', error: undefined } : i));
    void processEmailQueue();
  };

  const emailQueueActive = emailQueue.filter(i => i.status === 'queued' || i.status === 'sending');
  const emailQueueFailed = emailQueue.filter(i => i.status === 'failed');
  const emailQueueSent = emailQueue.filter(i => i.status === 'sent');

  React.useEffect(() => {
    if (emailQueueActive.length === 0 && emailQueueSent.length > 0) {
      const t = setTimeout(() => {
        updateEmailQueue(prev => prev.filter(i => i.status !== 'sent'));
      }, EMAIL_QUEUE_CLEAR_DELAY_MS);
      return () => clearTimeout(t);
    }
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [emailQueueActive.length, emailQueueSent.length]);

  const handleCopyKey = () => {
    navigator.clipboard.writeText(HARDCODED_KEY);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleFileUpload = (file: File) => {
    setErrorMsg('');
    setSuccessMsg('');
    setFilename(file.name);

    if (file.name.endsWith('.csv')) {
      Papa.parse(file, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          if (results.errors.length > 0) {
            setErrorMsg('Failed to parse CSV file.');
            return;
          }
          processParsedData(results.data as string[][]);
        }
      });
    } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'binary' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as string[][];
          processParsedData(json);
        } catch {
          setErrorMsg('Failed to parse Excel file.');
        }
      };
      reader.readAsBinaryString(file);
    } else {
      setErrorMsg('Unsupported file type. Please upload a CSV or Excel file.');
    }
  };

  const processParsedData = (data: string[][]) => {
    if (!data || data.length < 2) {
      setErrorMsg('File appears to be empty or missing headers.');
      return;
    }

    const headers = data[0].map(h => String(h).trim());
    setRawData(data);

    // Column detection is case-insensitive throughout — spreadsheet column names
    // have shifted between versions before and may shift again.
    const exactName        = headers.findIndex(h => h.toLowerCase() === 'clientname');
    const exactNum         = headers.findIndex(h => h.toLowerCase() === 'regno');
    const exactStatus      = headers.findIndex(h => h.toLowerCase() === 'status');
    const exactContactName = headers.findIndex(h => h.toLowerCase() === 'contactname');
    const exactEmail       = headers.findIndex(h => h.toLowerCase() === 'email');
    const exactMret        = headers.findIndex(h => h.toLowerCase() === 'mret');

    let foundName = exactName !== -1 ? exactName : -1;
    let foundNum  = exactNum  !== -1 ? exactNum  : -1;
    let foundStatus = exactStatus !== -1 ? exactStatus : -1;

    if (foundName === -1 || foundNum === -1) {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i].toLowerCase();
        if (foundName === -1 && /clientname|client.?name|company.?name/i.test(h)) foundName = i;
        if (foundNum  === -1 && /regno|company.?no|crn|reg.?no|number/i.test(h))  foundNum  = i;
      }
    }

    setNameColIdx(foundName !== -1 ? String(foundName) : '0');
    setNumberColIdx(foundNum !== -1 ? String(foundNum) : '1');
    setStatusColIdx(foundStatus);
    setContactNameColIdx(exactContactName);
    setEmailColIdx(exactEmail);
    setMretColIdx(exactMret);
    setPhase('mapped');
    const statusNote = foundStatus !== -1 ? ' Status filter active (non-Active rows will be skipped).' : '';
    setSuccessMsg(`Loaded ${data.length - 1} rows from spreadsheet.${statusNote}`);
  };

  const startLookup = async () => {
    if (!nameColIdx || !numberColIdx) {
      setErrorMsg('Please select both name and number columns.');
      return;
    }

    const nIdx = parseInt(nameColIdx, 10);
    const numIdx = parseInt(numberColIdx, 10);

    cancelledRef.current = false;
    setIsRunning(true);
    setPhase('running');
    setSuccessMsg('');
    setErrorMsg('');

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const madeUpToCutoff = madeUpToDate ? new Date(madeUpToDate) : null;
    if (madeUpToCutoff) madeUpToCutoff.setHours(23, 59, 59);

    const rows = rawData.slice(1).filter(r => r[nIdx] && String(r[nIdx]).trim());
    const results: ClientResult[] = [];
    const mismatchResults: MismatchResult[] = [];
    let noNum = 0, notActiveStatus = 0, notActiveCH = 0, apiCallCount = 0;

    // Pre-count rows that will actually hit the API, for the progress bar
    const eligibleCount = rows.filter(r => {
      const num = r[numIdx];
      if (!num || !String(num).trim()) return false;
      if (statusColIdx !== -1 && String(r[statusColIdx] || '').trim().toLowerCase() !== 'active') return false;
      return true;
    }).length;
    setProgressTotal(eligibleCount);

    for (let i = 0; i < rows.length; i++) {
      if (cancelledRef.current) break;
      const clientName = String(rows[i][nIdx] || 'Unknown').trim();
      const contactName = contactNameColIdx !== -1 ? String(rows[i][contactNameColIdx] || '').trim() : '';
      const email = emailColIdx !== -1 ? String(rows[i][emailColIdx] || '').trim() : '';
      const mret = mretColIdx !== -1 ? String(rows[i][mretColIdx] || '').trim() : '';
      const ourStatus = statusColIdx !== -1 ? String(rows[i][statusColIdx] || '').trim() : '';
      const rawNum = rows[i][numIdx];

      if (!rawNum || !String(rawNum).trim()) {
        noNum++;
        continue;
      }

      // Skip rows where Status is not Active (hard filter, before any API call)
      if (statusColIdx !== -1) {
        const rowStatus = String(rows[i][statusColIdx] || '').trim().toLowerCase();
        if (rowStatus !== 'active') {
          notActiveStatus++;
          continue;
        }
      }

      apiCallCount++;
      setProgressCurrent(apiCallCount);
      setProgressLabel(clientName);

      const result = await fetchCompany(String(rawNum));

      if (result.error) {
        results.push({ name: clientName, number: result.padded, dueDate: null, dueObj: null, madeUpTo: null, madeUpToObj: null, diffDays: null, status: 'error', error: result.error, contactName, email, mret });
      } else {
        const companyStatus = (result.data as Record<string, string>)?.company_status;
        const accountsData = (result.data as Record<string, Record<string, unknown>>)?.accounts;
        const dueStr = accountsData?.next_due as string | undefined;
        const madeUpToStr = accountsData?.next_made_up_to as string | undefined;
        const madeUpToObj = madeUpToStr ? new Date(madeUpToStr) : null;
        const nextAccounts = accountsData?.next_accounts as Record<string, string> | undefined;
        const periodStart = nextAccounts?.period_start_on;
        const periodEnd = nextAccounts?.period_end_on;

        if (companyStatus !== 'active') {
          // Still respect the made-up-to filter — only surface dissolved/inactive
          // companies whose outstanding period is actually in range. If we can't
          // tell (no made-up-to date returned), show it rather than hide it.
          const inRange = !madeUpToObj || !madeUpToCutoff || madeUpToObj <= madeUpToCutoff;
          if (inRange) {
            notActiveCH++;
            mismatchResults.push({ name: clientName, number: result.padded, ourStatus: ourStatus || 'Active', chStatus: companyStatus || 'unknown' });
          }
          if (i < rows.length - 1 && !cancelledRef.current) await delay(600);
          continue;
        }
        if (!dueStr) {
          results.push({ name: clientName, number: result.padded, dueDate: null, dueObj: null, madeUpTo: madeUpToStr || null, madeUpToObj, diffDays: null, status: 'error', error: 'No accounts due date returned', contactName, email, mret });
        } else {
          const dueObj = new Date(dueStr);
          dueObj.setHours(0, 0, 0, 0);
          const diffDays = Math.round((dueObj.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
          const status: ResultStatus = dueObj < today ? 'overdue' : diffDays <= AMBER_THRESHOLD_DAYS ? 'soon' : 'clear';
          results.push({ name: clientName, number: result.padded, dueDate: dueStr, dueObj, madeUpTo: madeUpToStr || null, madeUpToObj, diffDays, status, contactName, email, mret, periodStart, periodEnd });
        }
      }
      if (i < rows.length - 1 && !cancelledRef.current) await delay(600);
    }

    setAllResults(results);
    setMismatches(mismatchResults);
    setMismatchOpen(false);
    setPhase('done');
    setIsRunning(false);

    if (cancelledRef.current) {
      setSuccessMsg(`Lookup cancelled. Processed ${results.length} companies.`);
    } else {
      const skipParts = [];
      if (noNum > 0) skipParts.push(`${noNum} with no company number`);
      if (notActiveStatus > 0) skipParts.push(`${notActiveStatus} not Active in spreadsheet`);
      const skipStr = skipParts.length > 0 ? ` Skipped: ${skipParts.join(', ')}.` : '';
      const mismatchStr = notActiveCH > 0 ? ` ${notActiveCH} mismatch${notActiveCH === 1 ? '' : 'es'} found (see below).` : '';
      setSuccessMsg(`Lookup complete — ${results.length} companies checked.${skipStr}${mismatchStr}`);
    }
  };

  const testSingle = async () => {
    if (!nameColIdx || !numberColIdx || rawData.length < 2) return;
    setSuccessMsg('');
    setErrorMsg('');

    const nIdx = parseInt(nameColIdx, 10);
    const numIdx = parseInt(numberColIdx, 10);
    const row = rawData[1];

    const rawNum = row[numIdx];
    const clientName = String(row[nIdx] || 'Unknown').trim();

    if (!rawNum || !String(rawNum).trim()) {
      setErrorMsg(`Test row (${clientName}) has no company number.`);
      return;
    }

    const result = await fetchCompany(String(rawNum));
    if (result.error) {
      setErrorMsg(`Test failed for ${clientName} (${result.padded}): ${result.error}`);
    } else {
      const companyStatus = (result.data as Record<string, string>)?.company_status;
      if (companyStatus !== 'active') {
        setErrorMsg(`Test succeeded for ${clientName} (${result.padded}), but company is not active (${companyStatus}).`);
        return;
      }
      const accountsData = (result.data as Record<string, Record<string, string>>)?.accounts;
      const dueStr = accountsData?.next_due;
      if (!dueStr) {
        setErrorMsg(`Test succeeded for ${clientName} (${result.padded}), but no accounts due date was found.`);
      } else {
        setSuccessMsg(`Test succeeded! ${clientName} (${result.padded}) is active, next accounts due: ${dueStr}`);
      }
    }
  };

  const resetToMapped = () => {
    setAllResults([]);
    setMismatches([]);
    setMismatchOpen(false);
    setSuccessMsg('');
    setErrorMsg('');
    setSearch('');
    setCurrentFilter('all');
    setPhase('mapped');
  };

  const resetToSetup = () => {
    setPhase('setup');
    setFilename('');
    setRawData([]);
    setAllResults([]);
    setMismatches([]);
    setMismatchOpen(false);
    emailQueueRef.current = [];
    setEmailQueue([]);
    setSuccessMsg('');
    setErrorMsg('');
    setSearch('');
    setCurrentFilter('all');
    setStatusColIdx(-1);
    setContactNameColIdx(-1);
    setEmailColIdx(-1);
    setMretColIdx(-1);
  };

  const exportCSV = () => {
    const rows: string[][] = [['Client name', 'Company number', 'Made-up-to date', 'Due date', 'Days', 'Status', 'MRet', 'Notes']];
    filteredResults.forEach(r => rows.push([
      r.name, r.number, r.madeUpTo || '', r.dueDate || '',
      r.diffDays != null ? String(r.diffDays) : '',
      r.status,
      isRetainer(r.mret) ? 'Yes' : '',
      r.error || ''
    ]));
    rows.push([]);
    rows.push([`--- Companies House Mismatches (${mismatches.length}) ---`, '', '', '', '', '', '', '']);
    if (mismatches.length === 0) {
      rows.push(['None in this run', '', '', '', '', '', '', '']);
    } else {
      rows.push(['Client name', 'Company number', 'Our system status', 'Companies House status', 'Action', '', '', '']);
      mismatches.forEach(m => rows.push([m.name, m.number, m.ourStatus, m.chStatus, 'Review required', '', '', '']));
    }
    const csv = rows.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `accounts-due-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
  };

  const openEmailModal = (r: ClientResult) => {
    const overdueSuffix = r.status === 'overdue' ? ' - Overdue' : '';
    setModalRow(r);
    setDraftTo(r.email || '');
    setDraftSubject(`${r.name} Accounts documents required (ACT)${overdueSuffix}`);
    setDraftBody(buildEmailBody(r));
    setEmailModalOpen(true);
  };

  const handleSendEmail = () => {
    if (!modalRow) return;
    // Convert the plain-text draft to HTML so Make.com preserves line breaks and spacing —
    // plain text renders as an unstyled wall of text even with the module set to HTML.
    // £120 is bolded only in this final HTML output, not in the editable draft, so the
    // Textarea stays plain and reliably editable — this only re-bolds if "£120" still
    // appears verbatim, so a reworded/changed amount won't come through bold.
    const bodyHtml = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#333">'
      + draftBody
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/£120/g, '<strong>£120</strong>')
          .replace(/\n/g, '<br>')
      + '</div>';
    enqueueEmail({
      id: `${modalRow.number}-${Date.now()}`,
      number: modalRow.number,
      name: modalRow.name,
      to: draftTo,
      subject: draftSubject,
      bodyHtml,
    });
    setEmailModalOpen(false);
  };

  const statusOrder: Record<ResultStatus, number> = { overdue: 0, soon: 1, clear: 2, error: 3 };

  const filteredResults = useMemo(() => {
    const cutoff = madeUpToDate ? new Date(madeUpToDate) : null;
    if (cutoff) cutoff.setHours(23, 59, 59);

    return allResults
      .filter(r => {
        if (search && !r.name.toLowerCase().includes(search.toLowerCase()) && !r.number.includes(search)) return false;
        if (currentFilter !== 'all' && r.status !== currentFilter) return false;
        // Cumulative: any client whose made-up-to date is on or before the cutoff appears,
        // however far back — nothing drops off just because time passes.
        if (cutoff && r.madeUpToObj && r.madeUpToObj > cutoff) return false;
        return true;
      })
      .sort((a, b) => statusOrder[a.status] - statusOrder[b.status] || (a.diffDays ?? 9999) - (b.diffDays ?? 9999));
  }, [allResults, currentFilter, search, madeUpToDate]);

  const stats = useMemo(() => {
    let overdue = 0, soon = 0, clear = 0;
    allResults.forEach(r => {
      if (r.status === 'overdue') overdue++;
      if (r.status === 'soon') soon++;
      if (r.status === 'clear') clear++;
    });
    return { total: allResults.length, overdue, soon, clear };
  }, [allResults]);

  return (
    <div className="min-h-screen bg-slate-50 pb-20" data-print-date={format(new Date(), 'dd MMM yyyy')}>
      <header className="bg-white border-b border-slate-200 px-6 py-3 sticky top-0 z-10 shadow-sm print:hidden">
        <div className="max-w-6xl mx-auto flex items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <img src="/al-logo.png" alt="AL Accounting Solutions" className="h-10 w-auto" />
            <div className="border-l border-slate-200 pl-4">
              <h1 className="text-base font-semibold tracking-tight" style={{ color: '#176482' }}>Accounts Due Tracker</h1>
              <p className="text-xs text-slate-400 font-medium">Internal Use Only</p>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8 space-y-6">

        {successMsg && (
          <Alert className="bg-green-50 border-green-200 text-green-800 print:hidden">
            <CheckCircle2 className="h-4 w-4" color="currentColor" />
            <AlertTitle>Success</AlertTitle>
            <AlertDescription>{successMsg}</AlertDescription>
          </Alert>
        )}

        {errorMsg && (
          <Alert variant="destructive" className="print:hidden">
            <AlertCircle className="h-4 w-4" />
            <AlertTitle>Error</AlertTitle>
            <AlertDescription>{errorMsg}</AlertDescription>
          </Alert>
        )}

        {(phase === 'setup' || phase === 'mapped') && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <Card>
              <CardHeader className="pb-3">
                <CardTitle className="text-base font-semibold">Settings</CardTitle>
                <CardDescription>Configure API access and cutoff dates.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="api-key" className="text-slate-700">Companies House API Key</Label>
                  <div className="flex gap-2">
                    <Input
                      id="api-key"
                      data-testid="api-key-input"
                      value={HARDCODED_KEY}
                      readOnly
                      className="font-mono text-sm bg-slate-50"
                    />
                    <Button
                      variant="outline"
                      size="icon"
                      onClick={handleCopyKey}
                      data-testid="copy-btn"
                      className="shrink-0"
                    >
                      {copied ? <Check className="h-4 w-4 text-green-600" /> : <Copy className="h-4 w-4" />}
                    </Button>
                  </div>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="made-up-to" className="text-slate-700">Made up to date</Label>
                  <Input
                    id="made-up-to"
                    type="date"
                    data-testid="made-up-to-input"
                    value={madeUpToDate}
                    onChange={e => setMadeUpToDate(e.target.value)}
                  />
                  <p className="text-xs text-slate-500">Show all clients with accounts made up to this date or earlier. Run this for the end of the month you are working on.</p>
                </div>
              </CardContent>
            </Card>

            <Card>
              <CardHeader className="pb-3">
                <CardTitle className="text-base font-semibold">Data Source</CardTitle>
                <CardDescription>Upload client list to begin lookup.</CardDescription>
              </CardHeader>
              <CardContent>
                {!filename ? (
                  <div
                    className="border-2 border-dashed border-slate-300 rounded-lg p-8 flex flex-col items-center justify-center text-center cursor-pointer hover:bg-slate-50 transition-colors"
                    onClick={() => document.getElementById('file-upload')?.click()}
                    onDragOver={e => e.preventDefault()}
                    onDrop={e => {
                      e.preventDefault();
                      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                        handleFileUpload(e.dataTransfer.files[0]);
                      }
                    }}
                    data-testid="upload-zone"
                  >
                    <UploadCloud className="h-10 w-10 text-slate-400 mb-3" />
                    <p className="text-sm font-medium text-slate-700">Click or drag spreadsheet here</p>
                    <p className="text-xs text-slate-500 mt-1">Supports .csv, .xlsx, .xls</p>
                    <input
                      id="file-upload"
                      type="file"
                      className="hidden"
                      accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                      onChange={e => {
                        if (e.target.files && e.target.files[0]) {
                          handleFileUpload(e.target.files[0]);
                        }
                      }}
                      data-testid="file-input"
                    />
                  </div>
                ) : (
                  <div className="bg-slate-50 rounded-lg p-4 flex items-center justify-between border border-slate-200">
                    <div className="flex items-center gap-3">
                      <FileSpreadsheet className="h-8 w-8 text-blue-600" />
                      <div>
                        <p className="text-sm font-medium text-slate-900 truncate max-w-[200px]">{filename}</p>
                        <p className="text-xs text-slate-500">{rawData.length > 0 ? `${rawData.length - 1} rows` : 'Loading...'}</p>
                      </div>
                    </div>
                    <Button variant="ghost" size="sm" onClick={resetToSetup} className="text-slate-500">
                      Change
                    </Button>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        )}

        {phase === 'mapped' && rawData.length > 0 && (
          <Card>
            <CardHeader className="pb-3 border-b border-slate-100">
              <div className="flex items-center justify-between">
                <CardTitle className="text-base font-semibold">Column Mapping</CardTitle>
                <div className="flex items-center gap-2 text-xs text-slate-500">
                  <FileSpreadsheet className="h-3.5 w-3.5 shrink-0" />
                  <span className="truncate max-w-[180px]">{filename}</span>
                  <button onClick={resetToSetup} className="text-blue-600 hover:underline shrink-0">Change file</button>
                </div>
              </div>
            </CardHeader>
            <CardContent className="pt-6">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 mb-6">
                <div className="space-y-2">
                  <Label className="text-slate-700">Client name column</Label>
                  <Select value={nameColIdx} onValueChange={setNameColIdx}>
                    <SelectTrigger data-testid="name-col-select">
                      <SelectValue placeholder="Select column" />
                    </SelectTrigger>
                    <SelectContent>
                      {rawData[0].map((h, i) => (
                        <SelectItem key={i} value={String(i)}>{h || `Column ${i + 1}`}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label className="text-slate-700">Company number column</Label>
                  <Select value={numberColIdx} onValueChange={setNumberColIdx}>
                    <SelectTrigger data-testid="number-col-select">
                      <SelectValue placeholder="Select column" />
                    </SelectTrigger>
                    <SelectContent>
                      {rawData[0].map((h, i) => (
                        <SelectItem key={i} value={String(i)}>{h || `Column ${i + 1}`}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>

              <div className="flex gap-3">
                <Button onClick={startLookup} disabled={isRunning} data-testid="run-btn" className="flex-1">
                  <Play className="h-4 w-4 mr-2" />
                  Run lookup
                </Button>
                <Button variant="secondary" onClick={testSingle} disabled={isRunning} data-testid="test-btn">
                  Test single call
                </Button>
              </div>
            </CardContent>
          </Card>
        )}

        {phase === 'running' && (
          <Card className="border-blue-200 bg-blue-50/50">
            <CardContent className="pt-6">
              <div className="flex justify-between items-end mb-2">
                <div>
                  <h3 className="text-sm font-semibold text-blue-900 mb-1">Lookup in progress</h3>
                  <p className="text-xs text-blue-700">
                    {progressCurrent} of {progressTotal} — {progressLabel}
                  </p>
                </div>
                <Button
                  variant="destructive"
                  size="sm"
                  onClick={() => cancelledRef.current = true}
                  data-testid="cancel-btn"
                >
                  Cancel
                </Button>
              </div>
              <Progress value={progressTotal > 0 ? (progressCurrent / progressTotal) * 100 : 0} className="h-2" />
            </CardContent>
          </Card>
        )}

        {phase === 'done' && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">

            <div className="flex items-center justify-between print:hidden">
              <Button variant="outline" size="sm" onClick={resetToMapped} data-testid="back-btn">
                <ArrowLeft className="h-4 w-4 mr-2" />
                New lookup
              </Button>

              {emailQueue.length > 0 && (
                <div className={`flex items-center gap-2 rounded-full px-3 py-1.5 text-sm font-medium ${
                  emailQueueActive.length > 0
                    ? 'bg-blue-50 text-blue-700 border border-blue-200'
                    : emailQueueFailed.length > 0
                    ? 'bg-red-50 text-red-700 border border-red-200'
                    : 'bg-green-50 text-green-700 border border-green-200'
                }`}>
                  {emailQueueActive.length > 0 ? (
                    <>
                      <Loader2 className="h-3.5 w-3.5 animate-spin" />
                      Sending {emailQueueSent.length + emailQueueFailed.length + 1} of {emailQueue.length}…
                    </>
                  ) : emailQueueFailed.length > 0 ? (
                    <>
                      <AlertTriangle className="h-3.5 w-3.5" />
                      {emailQueueFailed.length} email{emailQueueFailed.length === 1 ? '' : 's'} failed — retry from the row{emailQueueFailed.length === 1 ? '' : 's'} below
                    </>
                  ) : (
                    <>
                      <CheckCircle2 className="h-3.5 w-3.5" />
                      All emails sent
                    </>
                  )}
                </div>
              )}
            </div>

            {madeUpToDate && (
              <p className="text-lg font-bold text-slate-800">
                Showing clients with accounts made up to {format(new Date(madeUpToDate), 'dd MMM yyyy')} or earlier
              </p>
            )}

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <Card>
                <CardContent className="p-4 flex flex-col justify-center">
                  <p className="text-sm font-medium text-slate-500">Total Clients</p>
                  <p className="text-3xl font-bold text-slate-900">{stats.total}</p>
                </CardContent>
              </Card>
              <Card className="border-red-200 bg-red-50/30">
                <CardContent className="p-4 flex flex-col justify-center">
                  <p className="text-sm font-medium text-red-600">Overdue</p>
                  <p className="text-3xl font-bold text-red-700">{stats.overdue}</p>
                </CardContent>
              </Card>
              <Card className="border-amber-200 bg-amber-50/30">
                <CardContent className="p-4 flex flex-col justify-center">
                  <p className="text-sm font-medium text-amber-600">Due within {AMBER_THRESHOLD_DAYS} days</p>
                  <p className="text-3xl font-bold text-amber-700">{stats.soon}</p>
                </CardContent>
              </Card>
              <Card className="border-green-200 bg-green-50/30">
                <CardContent className="p-4 flex flex-col justify-center">
                  <p className="text-sm font-medium text-green-600">Clear</p>
                  <p className="text-3xl font-bold text-green-700">{stats.clear}</p>
                </CardContent>
              </Card>
            </div>

            <Card>
              <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 print:hidden">
                <div className="flex items-center gap-2 overflow-x-auto pb-2 sm:pb-0">
                  <Button
                    variant={currentFilter === 'all' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => setCurrentFilter('all')}
                    data-testid="filter-all"
                  >
                    All
                  </Button>
                  <Button
                    variant={currentFilter === 'overdue' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => setCurrentFilter('overdue')}
                    className={currentFilter === 'overdue' ? "bg-red-600 hover:bg-red-700 text-white border-red-600" : "text-red-600 border-red-200 hover:bg-red-50"}
                    data-testid="filter-overdue"
                  >
                    Overdue
                  </Button>
                  <Button
                    variant={currentFilter === 'soon' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => setCurrentFilter('soon')}
                    className={currentFilter === 'soon' ? "bg-amber-600 hover:bg-amber-700 text-white border-amber-600" : "text-amber-600 border-amber-200 hover:bg-amber-50"}
                    data-testid="filter-soon"
                  >
                    Due soon
                  </Button>
                  <Button
                    variant={currentFilter === 'clear' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => setCurrentFilter('clear')}
                    className={currentFilter === 'clear' ? "bg-green-600 hover:bg-green-700 text-white border-green-600" : "text-green-600 border-green-200 hover:bg-green-50"}
                    data-testid="filter-clear"
                  >
                    Clear
                  </Button>
                  <Button
                    variant={currentFilter === 'error' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => setCurrentFilter('error')}
                    className={currentFilter === 'error' ? "bg-slate-600 hover:bg-slate-700 text-white border-slate-600" : "text-slate-600 border-slate-200 hover:bg-slate-50"}
                    data-testid="filter-error"
                  >
                    Errors
                  </Button>
                </div>

                <div className="flex items-center gap-2">
                  <div className="relative">
                    <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-slate-400" />
                    <Input
                      placeholder="Search company..."
                      value={search}
                      onChange={e => setSearch(e.target.value)}
                      className="pl-9 w-[200px] h-9"
                      data-testid="search-input"
                    />
                  </div>
                  <Button variant="outline" size="sm" onClick={exportCSV} data-testid="export-btn" className="h-9 shrink-0">
                    <Download className="h-4 w-4 mr-2" /> Export
                  </Button>
                  <Button variant="outline" size="sm" onClick={() => window.print()} className="h-9 shrink-0">
                    <Printer className="h-4 w-4 mr-2" /> Print
                  </Button>
                </div>
              </div>

              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow className="hover:bg-transparent">
                      <TableHead className="w-[280px]">Client name</TableHead>
                      <TableHead>Company number</TableHead>
                      <TableHead>Made-up-to date</TableHead>
                      <TableHead>Due date</TableHead>
                      <TableHead>Days</TableHead>
                      <TableHead className="text-right">Status</TableHead>
                      <TableHead className="text-center">MRet</TableHead>
                      <TableHead className="text-right w-[140px] print:hidden">Actions</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {filteredResults.length === 0 ? (
                      <TableRow>
                        <TableCell colSpan={8} className="h-32 text-center text-slate-500">
                          No results found matching criteria.
                        </TableCell>
                      </TableRow>
                    ) : (
                      filteredResults.map((r, i) => (
                        <TableRow key={i}>
                          <TableCell className="font-medium text-slate-900">{r.name}</TableCell>
                          <TableCell className="font-mono text-sm text-slate-600">{r.number}</TableCell>
                          <TableCell className="text-slate-600">
                            {r.madeUpTo ? format(new Date(r.madeUpTo), 'dd MMM yyyy') : '-'}
                          </TableCell>
                          <TableCell className="text-slate-600">
                            {r.dueDate ? format(new Date(r.dueDate), 'dd MMM yyyy') : '-'}
                          </TableCell>
                          <TableCell>
                            {r.status === 'overdue' && <span className="font-bold text-red-600">{Math.abs(r.diffDays!)} days ago</span>}
                            {r.status === 'soon' && <span className="font-bold text-amber-600">{r.diffDays} days</span>}
                            {r.status === 'clear' && <span className="text-slate-500">{r.diffDays} days</span>}
                            {r.status === 'error' && <span className="text-xs text-red-500 leading-tight">{r.error || 'Unknown error'}</span>}
                          </TableCell>
                          <TableCell className="text-right">
                            {r.status === 'overdue' && <Badge variant="destructive" className="bg-red-100 text-red-800 hover:bg-red-100">Overdue</Badge>}
                            {r.status === 'soon' && <Badge variant="outline" className="bg-amber-100 border-amber-200 text-amber-800">Due soon</Badge>}
                            {r.status === 'clear' && <Badge variant="outline" className="bg-green-100 border-green-200 text-green-800">Clear</Badge>}
                            {r.status === 'error' && <Badge variant="secondary" className="bg-slate-100 text-slate-500">Error</Badge>}
                          </TableCell>
                          <TableCell className="text-center text-sm text-slate-600">
                            {isRetainer(r.mret) ? (
                              <Badge variant="outline" className="bg-teal-50 border-teal-200 text-teal-700 text-xs font-medium">
                                Monthly Retainer
                              </Badge>
                            ) : ''}
                          </TableCell>
                          <TableCell className="text-right print:hidden">
                            {(() => {
                              const queueItem = emailQueue.find(q => q.number === r.number);
                              if (queueItem?.status === 'queued' || queueItem?.status === 'sending') {
                                return (
                                  <Badge variant="outline" className="h-7 px-2.5 text-xs gap-1 bg-blue-50 border-blue-200 text-blue-700 font-normal">
                                    <Loader2 className="h-3 w-3 animate-spin" />
                                    {queueItem.status === 'sending' ? 'Sending…' : 'Queued'}
                                  </Badge>
                                );
                              }
                              if (queueItem?.status === 'sent') {
                                return (
                                  <Badge variant="outline" className="h-7 px-2.5 text-xs gap-1 bg-green-50 border-green-200 text-green-700 font-normal">
                                    <CheckCircle2 className="h-3 w-3" /> Sent
                                  </Badge>
                                );
                              }
                              if (queueItem?.status === 'failed') {
                                return (
                                  <div className="flex items-center justify-end gap-1.5">
                                    <Badge variant="destructive" className="h-7 px-2.5 text-xs gap-1 bg-red-100 text-red-800 hover:bg-red-100 font-normal" title={queueItem.error}>
                                      <XCircle className="h-3 w-3" /> Failed
                                    </Badge>
                                    <Button
                                      size="sm"
                                      variant="outline"
                                      className="h-7 text-xs gap-1 border-red-300 text-red-700 hover:bg-red-50"
                                      onClick={() => retryEmail(queueItem.id)}
                                    >
                                      <RefreshCw className="h-3 w-3" /> Retry
                                    </Button>
                                  </div>
                                );
                              }
                              return isRetainer(r.mret) ? (
                                <Button
                                  size="sm"
                                  variant="outline"
                                  className="h-7 text-xs gap-1 border-slate-300 text-slate-500 hover:bg-slate-50"
                                  onClick={() => openEmailModal(r)}
                                >
                                  Docs only
                                </Button>
                              ) : (
                                <Button
                                  size="sm"
                                  variant="outline"
                                  className="h-7 text-xs gap-1 border-slate-300 hover:bg-blue-50 hover:border-blue-300 hover:text-blue-700"
                                  onClick={() => openEmailModal(r)}
                                >
                                  <Mail className="h-3 w-3" /> Email
                                </Button>
                              );
                            })()}
                          </TableCell>
                        </TableRow>
                      ))
                    )}
                  </TableBody>
                </Table>
              </div>
            </Card>

            <div className="border border-orange-200 rounded-lg bg-orange-50/40">
              <button
                className="w-full flex items-center justify-between px-5 py-4 text-left print:hidden"
                onClick={() => setMismatchOpen(o => !o)}
              >
                <div className="flex items-center gap-2">
                  <AlertTriangle className="h-4 w-4 text-orange-500 shrink-0" />
                  <span className="font-semibold text-orange-900">
                    Companies House Mismatches ({mismatches.length})
                  </span>
                </div>
                <ChevronDown className={`h-4 w-4 text-orange-600 transition-transform ${mismatchOpen ? 'rotate-180' : ''}`} />
              </button>
              <div className="hidden print:block px-5 pt-4">
                <span className="font-semibold text-orange-900">
                  Companies House Mismatches ({mismatches.length})
                </span>
              </div>

              <div className={`px-5 pb-5 space-y-3 ${mismatchOpen ? '' : 'hidden print:block'}`}>
                  <p className="text-sm text-orange-800 border-t border-orange-200 pt-4">
                    This report shows discrepancies between our internal records and Companies House. The following clients are marked as active in our system but are showing a different status on Companies House. Please review and update internal records accordingly.
                  </p>
                  {mismatches.length === 0 ? (
                    <p className="text-sm text-orange-700 italic">None in this run — every Active client checked out as active with Companies House.</p>
                  ) : (
                    <div className="overflow-x-auto rounded border border-orange-200 bg-white">
                      <Table>
                        <TableHeader>
                          <TableRow className="hover:bg-transparent bg-orange-50">
                            <TableHead className="w-[280px]">Client name</TableHead>
                            <TableHead>Company number</TableHead>
                            <TableHead>Our system status</TableHead>
                            <TableHead>Companies House status</TableHead>
                            <TableHead>Action</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {mismatches.map((m, i) => (
                            <TableRow key={i}>
                              <TableCell className="font-medium text-slate-900">{m.name}</TableCell>
                              <TableCell className="font-mono text-sm text-slate-600">{m.number}</TableCell>
                              <TableCell className="text-slate-600">{m.ourStatus}</TableCell>
                              <TableCell>
                                <Badge variant="outline" className="bg-orange-50 border-orange-300 text-orange-800 capitalize print:border-0 print:bg-transparent print:p-0">
                                  {m.chStatus.replace(/-/g, ' ')}
                                </Badge>
                              </TableCell>
                              <TableCell className="text-sm text-slate-500">Review required</TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  )}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Email draft modal */}
      <Dialog open={emailModalOpen} onOpenChange={setEmailModalOpen}>
        <DialogContent className="max-w-2xl">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Mail className="h-5 w-5 text-blue-600" />
              Draft email — {modalRow?.name}
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-4 py-2">
            <div className="space-y-1.5">
              <Label className="text-sm font-medium text-slate-700">To</Label>
              <Input
                value={draftTo}
                onChange={e => setDraftTo(e.target.value)}
                placeholder="recipient@example.com"
                className="font-mono text-sm"
              />
              <div className="flex items-center gap-3 mt-1">
                {!modalRow?.email && (
                  <p className="text-xs text-amber-600 flex items-center gap-1">
                    <AlertTriangle className="h-3 w-3" /> No address in spreadsheet — enter manually.
                  </p>
                )}
                <button
                  type="button"
                  className="text-xs text-blue-600 hover:underline ml-auto shrink-0"
                  onClick={() => setDraftTo(TEST_EMAIL)}
                >
                  Use test address
                </button>
              </div>
            </div>

            <div className="space-y-1.5">
              <Label className="text-sm font-medium text-slate-700">Subject</Label>
              <Input
                value={draftSubject}
                onChange={e => setDraftSubject(e.target.value)}
                className="text-sm"
              />
            </div>

            <div className="space-y-1.5">
              <Label className="text-sm font-medium text-slate-700">Body</Label>
              <Textarea
                value={draftBody}
                onChange={e => setDraftBody(e.target.value)}
                rows={18}
                className="text-sm resize-none leading-relaxed"
              />
            </div>
          </div>

          <DialogFooter className="gap-2">
            <Button variant="outline" onClick={() => setEmailModalOpen(false)}>
              Cancel
            </Button>
            <Button
              onClick={handleSendEmail}
              disabled={!draftTo.trim()}
              className="gap-2 bg-blue-600 hover:bg-blue-700 text-white"
            >
              <Send className="h-4 w-4" /> Send email
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Toaster richColors position="bottom-right" />
    </div>
  );
}
