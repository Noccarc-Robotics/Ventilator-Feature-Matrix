import React, { useEffect, useMemo, useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Switch } from "@/components/ui/switch";
import { Badge } from "@/components/ui/badge";
import { Download, Upload, Plus, Trash2, RefreshCw, Edit3, Eraser, Filter, FileSpreadsheet } from "lucide-react";
import { motion } from "framer-motion";
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as XLSX from "xlsx";

/**
 * Ventilator Feature Matrix – Single-file React app (Excel-backed)
 * - Color-coded (Enabled = green, Disabled = red, Blank/NA = gray)
 * - Filter by Mode, Patient Category, and Mode Type (Invasive/Non-Invasive)
 * - Add/Edit features with notes
 * - Built-in dataset parsed from your Excel so you DON'T need to import
 * - Export to CSV/JSON/XLSX
 */

const LS_KEY = "ventilator-feature-matrix";
const ALL = "__ALL__"; // filters' All option
const NA = "__NA__";  // Enabled = N/A (maps to "")

export type Row = {
  featureGroup: string;
  featureName: string;
  patientCategory: string; // Adult, Pediatric, Neonatal
  modeType: string;        // Invasive | Non-Invasive
  mode: string;            // PC, VC, PRVC, SIMV+PC, etc.
  enabled: string;         // Y|N|""
  notes?: string;
};

// === Mode catalog restricted to your Excel header list ===
const MODE_CATALOG: Record<string, string[]> = {
  Invasive: [
    "PC", "VC", "PRVC", "SIMV+PC", "SIMV+VC", "SIMV+PRVC", "PS/CPAP", "VS-VC", "VS-PRVC", "APRV"
  ],
  "Non-Invasive": [
    "NIV-PS", "HFOT", "DualPap", "NIV-PC", "nCPAP"
  ],
};

// === Built-in dataset from your Excel (embedded JSON) ===
const INITIAL_ROWS_BASE64 = `W3siZmVhdHVyZUdyb3VwIjoiVmlld3MiLCJmZWF0dXJlTmFtZSI6IlRpbGUgVmlldyIsInBhdGllbnRDYXRlZ29yeSI6IkFkdWx0IiwibW9kZVR5cGUiOiJJbnZhc2l2ZSIsIm1vZGUiOiJQQyIsImVuYWJsZWQiOiJZIn0seyJmZWF0dXJlR3JvdXAiOiJWaWV3cyIsImZlYXR1cmVOYW1lIjoiVGlsZSBWaWV3IiwicGF0aWVudENhdGVnb3J5IjoiQWR1bHQiLCJtb2RlVHlwZSI6IkludmFzaXZlIiwibW9kZSI6IlZDIiwiZW5hYmxlZCI6IlkifSx7ImZlYXR1cmVHcm91cCI6IlZpZXdzIiwiZmVhdHVyZU5hbWUiOiJUaWxlIFZpZXciLCJwYXRpZW50Q2F0ZWdvcnkiOiJBZHVsdCIsIm1vZGVUeXBlIjoiSW52YXNpdmUiLCJtb2RlIjoiUFJWQyIsImVuYWJsZWQiOiJZIn0seyJmZWF0dXJlR3JvdXAiOiJWaWV3cyIsImZlYXR1cmVOYW1lIjoiVGlsZSBWaWV3IiwicGF0aWVudENhdGVnb3J5IjoiQWR1bHQiLCJtb2RlVHlwZSI6IkludmFzaXZlIiwibW9kZSI6IlNJTXYrUEMiLCJlbmFibGVkIjoiWSJ9LHsiZmVhdHVyZUdyb3VwIjoiVmlld3MiLCJmZWF0dXJlTmFtZSI6IlRpbGUgVmlldyIsInBhdGllbnRDYXRlZ29yeSI6IkFkdWx0IiwibW9kZVR5cGUiOiJJbnZhc2l2ZSIsIm1vZGUiOiJTSU1WK1ZDIiwiZW5hYmxlZCI6IlkifSx7ImZlYXR1cmVHcm91cCI6IlZpZXdzIiwiZmVhdHVyZU5hbWUiOiJUaWxlIFZpZXciLCJwYXRpZW50Q2F0ZWdvcnkiOiJBZHVsdCIsIm1vZGVUeXBlIjoiSW52YXNpdmUiLCJtb2RlIjoiU0lNVitQUlZDIiwiZW5hYmxlZCI6IlkifSx7ImZlYXR1cmVHcm91cCI6IlZpZXdzIiwiZmVhdHVyZU5hbWUiOiJUaWxlIFZpZXciLCJwYXRpZW50Q2F0ZWdvcnkiOiJBZHVsdCIsIm1vZGVUeXBlIjoiSW52YXNpdmUiLCJtb2RlIjoiUFMvQ1BBUCIsImVuYWJsZWQiOiJZIn0seyJmZWF0dXJlR3JvdXAiOiJWaWV3cyIsImZlYXR1cmVOYW1lIjoiVGlsZSBWaWV3IiwicGF0aWVudENhdGVnb3J5IjoiQWR1bHQiLCJtb2RlVHlwZSI6IkludmFzaXZlIiwibW9kZSI6IlZTLVZDIiwiZW5hYmxlZCI6IlkifSx7ImZlYXR1cmVHcm91cCI6IlZpZXdzIiwiZmVhdHVyZU5hbWUiOiJUaWxlIFZpZXciLCJwYXRpZW50Q2F0ZWdvcnkiOiJBZHVsdCIsIm1vZGVUeXBlIjoiSW52YXNpdmUiLCJtb2RlIjoiVlMtUFJWQyIsImVuYWJsZWQiOiJZIn0seyJmZWF0dXJlR3JvdXAiOiJWaWV3cyIsImZlYXR1cmVOYW1lIjoiVGlsZSBWaWV3IiwicGF0aWVudENhdGVnb3J5IjoiQWR1bHQiLCJtb2RlVHlwZSI6IkludmFzaXZlIiwibW9kZSI6IkFQUlYiLCJlbmFibGVkIjoiWSJ9LHsiZmVhdHVyZUdyb3VwIjoiVmlld3MiLCJmZWF0dXJlTmFtZSI6IlRpbGUgVmlldyIsInBhdGllbnRDYXRlZ29yeSI6IkFkdWx0IiwibW9kZVR5cGUiOiJOb24tSW52YXNpdmUiLCJtb2RlIjoiTklWLVBTIiwiZW5hYmxlZCI6IlkifSx7ImZlYXR1cmVHcm91cCI6IlZpZXdzIiwiZmVhdHVyZU5hbWUiOiJUaWxlIFZpZXciLCJwYXRpZW50Q2F0ZWdvcnkiOiJBZHVsdCIsIm1vZGVUeXBlIjoiTm9uLUludmFzaXZlIiwibW9kZSI6IkhGT1QiLCJlbmFibGVkIjoiWSJ9LHsiZmVhdHVyZUdyb3VwIjoiVmlld3MiLCJmZWF0dXJlTmFtZSI6IlRpbGUgVmlldyIsInBhdGllbnRDYXRlZ29yeSI6IkFkdWx0IiwibW9kZVR5cGUiOiJOb24tSW52YXNpdmUiLCJtb2RlIjoiRHVhbFBhcCIsImVuYWJsZWQiOiJZIn0seyJmZWF0dXJlR3JvdXAiOiJWaWV3cyIsImZlYXR1cmVOYW1lIjoiVGlsZSBWaWV3IiwicGF0aWVudENhdGVnb3J5IjoiQWR1bHQiLCJtb2RlVHlwZSI6Ik5vbi1JbnZhc2l2ZSIsIm1vZGUiOiJOSVYtUEMiLCJlbmFibGVkIjoiWSJ9LHsiZmVhdHVyZUdyb3VwIjoiVmlld3MiLCJmZWF0dXJlTmFtZSI6Ikh5YnJpZCBWaWV3IiwicGF0aWVudENhdGVnb3J5IjoiQWR1bHQiLCJtb2RlVHlwZSI6IkludmFzaXZlIiwibW9kZSI6IlBDIiwiZW5hYmxlZCI6IlkifV0`;
const INITIAL_ROWS: Row[] = JSON.parse(typeof atob !== 'undefined' ? atob(INITIAL_ROWS_BASE64) : (Buffer ? Buffer.from(INITIAL_ROWS_BASE64,'base64').toString('utf-8') : '[]'));

function loadInitial(): Row[] {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  return INITIAL_ROWS; // default to Excel-backed data
}

function save(rows: Row[]) { localStorage.setItem(LS_KEY, JSON.stringify(rows)); }

function badgeForEnabled(v: string) {
  const val = (v || "").trim().toUpperCase();
  if (val === "Y") return <Badge className="bg-green-600 hover:bg-green-600">Enabled</Badge>;
  if (val === "N") return <Badge className="bg-red-600 hover:bg-red-600">Disabled</Badge>;
  return <Badge className="bg-gray-500 hover:bg-gray-500">N/A</Badge>;
}

function uniqueSorted(rows: Row[], key: keyof Row) {
  return Array.from(new Set(rows.map(r => (r[key] || "").toString().trim()))).filter(Boolean).sort((a,b)=>a.localeCompare(b));
}

function normalizeModeType(s: string) {
  const t = (s || "").toLowerCase();
  if (/non\s*inv/.test(t)) return "Non-Invasive";
  if (/inv/.test(t)) return "Invasive";
  return s || "";
}
function normalizeMode(s: string) {
  const x = (s || "").toUpperCase().replace(/\s+/g, " ").trim();
  if (/^PCV$/.test(x)) return "PC";
  if (/^VCV$/.test(x)) return "VC";
  if (/^PRVC$/.test(x)) return "PRVC";
  if (/^(SIMV\s*\+?\s*PC|SIMV[-\s]*PC)$/i.test(x)) return "SIMV+PC";
  if (/^(SIMV\s*\+?\s*VC|SIMV[-\s]*VC)$/i.test(x)) return "SIMV+VC";
  if (/^(SIMV\s*\+?\s*PRVC|SIMV[-\s]*PRVC)$/i.test(x)) return "SIMV+PRVC";
  if (/^(PS\s*\/\s*CPAP|CPAP\s*\+\s*PS)$/i.test(x)) return "PS/CPAP";
  if (/^VS-?VC$/i.test(x)) return "VS-VC";
  if (/^VS-?PRVC$/i.test(x)) return "VS-PRVC";
  if (/^HFOT|^HFNC$/i.test(x)) return "HFOT";
  if (/^APRV$/i.test(x)) return "APRV";
  if (/^NIV[-\s]?PC$/i.test(x)) return "NIV-PC";
  if (/^NIV[-\s]?PS$/i.test(x) || /^NIV[-\s]?\(PS\)$/i.test(x) || /^Niv-PS$/i.test(x)) return "NIV-PS";
  if (/^NCPAP$/i.test(x) || /^N\s*CPAP$/i.test(x) || /^N-?CPAP$/i.test(x)) return "nCPAP";
  if (/^DUAL\s*PAP$/i.test(x) || /^DUALPAP$/i.test(x)) return "DualPap";
  return x;
}

export default function App() {
  const [rows, setRows] = useState<Row[]>(loadInitial());
  const [mode, setMode] = useState<string>("");
  const [patientCategory, setPatientCategory] = useState<string>("");
  const [modeType, setModeType] = useState<string>("");
  const [editable, setEditable] = useState<boolean>(false);

  useEffect(() => { save(rows); }, [rows]);

  const modesFromData = useMemo(() => uniqueSorted(rows, "mode"), [rows]);
  const categories = useMemo(() => uniqueSorted(rows, "patientCategory"), [rows]);
  const modeTypes = useMemo(() => uniqueSorted(rows, "modeType"), [rows]);

  const catalogForType = (modeType && MODE_CATALOG[modeType]) ? MODE_CATALOG[modeType] : Array.from(new Set(Object.values(MODE_CATALOG).flat()));
  const candidateModes = Array.from(new Set([...modesFromData, ...catalogForType])).sort((a,b)=>a.localeCompare(b));

  const filtered = useMemo(() => rows.filter(r =>
    (!mode || normalizeMode(r.mode) === mode) &&
    (!patientCategory || r.patientCategory === patientCategory) &&
    (!modeType || normalizeModeType(r.modeType) === modeType)
  ), [rows, mode, patientCategory, modeType]);

  function addRow(newRow: Row) {
    setRows(prev => [...prev, { ...newRow, mode: normalizeMode(newRow.mode), modeType: normalizeModeType(newRow.modeType) }]);
  }
  function updateRow(idx: number, patch: Partial<Row>) {
    const norm: Partial<Row> = { ...patch };
    if (patch.mode !== undefined) norm.mode = normalizeMode(patch.mode || "");
    if (patch.modeType !== undefined) norm.modeType = normalizeModeType(patch.modeType || "");
    setRows(prev => prev.map((r, i) => i === idx ? { ...r, ...norm } : r));
  }
  function deleteRow(idx: number) { setRows(prev => prev.filter((_, i) => i !== idx)); }
  function clearFilters() { setMode(""); setPatientCategory(""); setModeType(""); }

  function exportJSON() {
    const blob = new Blob([JSON.stringify(rows, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = "ventilator_features.json"; a.click(); URL.revokeObjectURL(url);
  }
  function exportCSV() {
    const header = ["featureGroup","featureName","patientCategory","modeType","mode","enabled","notes"]; 
    const body = rows.map(r => header.map(h => (r as any)[h] ?? ""));
    const csv = [header.join(","), ...body.map(line => line.map(v => `"${String(v).replace(/"/g,'""')}"`).join(","))].join("\n");
    const blob = new Blob([csv], { type: "text/csv" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = "ventilator_features.csv"; a.click(); URL.revokeObjectURL(url);
  }
  function exportXLSX() {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Features');
    XLSX.writeFile(wb, 'ventilator_features.xlsx');
  }

  const quickModes = ["PC","VC","PRVC","SIMV+PC","SIMV+VC","SIMV+PRVC","PS/CPAP","VS-VC","VS-PRVC","APRV","NIV-PS","HFOT","DualPap","NIV-PC","nCPAP"];

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-white p-4 md:p-8">
      <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}>
        <h1 className="text-2xl md:text-3xl font-semibold mb-2">Ventilator Feature Matrix</h1>
        <p className="text-slate-600 mb-6">Select Mode, Mode Type, and Patient Category to see color-coded availability pulled from your Excel.</p>
      </motion.div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
        <Card>
          <CardHeader className="pb-2"><CardTitle className="text-base flex items-center gap-2"><Filter className="w-4 h-4"/> Filters</CardTitle></CardHeader>
          <CardContent className="flex flex-col gap-3">
            <div className="flex gap-2">
              <Input placeholder="Search mode (e.g., PC, PRVC, NIV-PS, nCPAP)" value={mode} onChange={e=> setMode(normalizeMode(e.target.value))} />
            </div>
            <Select value={mode} onValueChange={(v)=> setMode(v === ALL ? "" : v)}>
              <SelectTrigger><SelectValue placeholder="Mode (e.g., PC, VC, PRVC)" /></SelectTrigger>
              <SelectContent>
                <SelectItem value={ALL}>All Modes</SelectItem>
                {candidateModes.map(m => <SelectItem key={m} value={m}>{m}</SelectItem>)}
              </SelectContent>
            </Select>
            <div className="flex flex-wrap gap-2">
              {quickModes.map(q => (
                <Button key={q} size="sm" variant={mode===q?"default":"outline"} onClick={()=>setMode(q)}>{q}</Button>
              ))}
            </div>

            <Select value={patientCategory} onValueChange={(v)=> setPatientCategory(v === ALL ? "" : v)}>
              <SelectTrigger><SelectValue placeholder="Patient Category" /></SelectTrigger>
              <SelectContent>
                <SelectItem value={ALL}>All Categories</SelectItem>
                {categories.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
              </SelectContent>
            </Select>
            <Select value={modeType} onValueChange={(v)=> setModeType(v === ALL ? "" : v)}>
              <SelectTrigger><SelectValue placeholder="Mode Type (Invasive / Non-Invasive)" /></SelectTrigger>
              <SelectContent>
                <SelectItem value={ALL}>Both</SelectItem>
                {Object.keys(MODE_CATALOG).map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
              </SelectContent>
            </Select>
            <div className="flex items-center gap-3 pt-1">
              <Switch id="editable" checked={editable} onCheckedChange={setEditable} />
              <label htmlFor="editable" className="text-sm text-slate-700 flex items-center gap-2"><Edit3 className="w-4 h-4"/> Inline Edit</label>
            </div>
            <div className="flex gap-2 pt-1">
              <Button variant="outline" onClick={clearFilters}><Eraser className="w-4 h-4 mr-2"/>Clear Filters</Button>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader className="pb-2"><CardTitle className="text-base">Export (built-in data)</CardTitle></CardHeader>
          <CardContent className="flex flex-col gap-3">
            <div className="flex gap-2">
              <Button variant="outline" onClick={exportCSV}><Download className="w-4 h-4 mr-2"/>Export CSV</Button>
              <Button variant="outline" onClick={exportJSON}><Download className="w-4 h-4 mr-2"/>Export JSON</Button>
              <Button variant="outline" onClick={exportXLSX}><FileSpreadsheet className="w-4 h-4 mr-2"/>Export Excel</Button>
              <Button variant="ghost" onClick={()=>setRows(INITIAL_ROWS)}><RefreshCw className="w-4 h-4 mr-2"/>Reset to Excel data</Button>
            </div>
            <p className="text-xs text-slate-500">Data is preloaded from your spreadsheet. You can still edit inline; it saves to your browser.</p>
          </CardContent>
        </Card>

        <AddFeatureCard onAdd={addRow} />
      </div>

      <Tabs defaultValue="selection">
        <TabsList>
          <TabsTrigger value="selection">Selected Mode View</TabsTrigger>
          <TabsTrigger value="table">Table</TabsTrigger>
          <TabsTrigger value="matrix">Matrix by Feature</TabsTrigger>
        </TabsList>
        <TabsContent value="selection">
          <SelectedModeView rows={rows} mode={mode} modeType={modeType} patientCategory={patientCategory} onAddQuick={(r)=>addRow(r)} />
        </TabsContent>
        <TabsContent value="table">
          <FeatureTable rows={filtered} editable={editable} onUpdate={updateRow} onDelete={deleteRow} />
        </TabsContent>
        <TabsContent value="matrix">
          <FeatureMatrix rows={filtered} />
        </TabsContent>
      </Tabs>

      <SelfTests rows={rows} />
    </div>
  );
}

function SelectedModeView({ rows, mode, modeType, patientCategory, onAddQuick }: { rows: Row[]; mode: string; modeType: string; patientCategory: string; onAddQuick: (r: Row)=>void; }) {
  const allFeatures = useMemo(() => {
    const m = new Map<string, {featureGroup:string, featureName:string}>();
    for (const r of rows) { const key = r.featureGroup+"|"+r.featureName; if (!m.has(key)) m.set(key, {featureGroup:r.featureGroup, featureName:r.featureName}); }
    return Array.from(m.values()).sort((a,b)=> a.featureGroup.localeCompare(b.featureGroup) || a.featureName.localeCompare(b.featureName));
  }, [rows]);

  const exact = useMemo(() => {
    if (!mode || !modeType || !patientCategory) return [] as Row[];
    return rows.filter(r => normalizeMode(r.mode)===mode && normalizeModeType(r.modeType)===modeType && r.patientCategory===patientCategory);
  }, [rows, mode, modeType, patientCategory]);

  const map = useMemo(() => { const m = new Map<string, Row>(); for (const r of exact) m.set(r.featureGroup+"|"+r.featureName, r); return m; }, [exact]);
  const haveSelection = !!mode && !!modeType && !!patientCategory;

  return (
    <Card>
      <CardHeader className="pb-2"><CardTitle className="text-base">Selected Mode View</CardTitle></CardHeader>
      <CardContent>
        {!haveSelection && (<div className="text-sm text-slate-600">Select <strong>Mode</strong>, <strong>Mode Type</strong> (Invasive / Non-Invasive), and <strong>Patient Category</strong> (Adult / Pediatric / Neonatal) above to see the feature list.</div>)}
        {haveSelection && (
          <div className="mb-3 text-sm">Showing <strong>{patientCategory}</strong> / <strong>{modeType}</strong> / <strong>{mode}</strong>
            <div className="text-slate-500">Features are color-coded. Use Add buttons to quickly fill missing entries for this exact combination.</div>
          </div>
        )}
        {haveSelection && (
          <div className="overflow-auto">
            <table className="min-w-full text-sm">
              <thead className="bg-slate-100"><tr><th className="text-left p-3">Feature Group</th><th className="text-left p-3">Feature</th><th className="text-left p-3">Enabled?</th><th className="text-left p-3">Notes</th><th className="text-left p-3">Quick</th></tr></thead>
              <tbody>
                {allFeatures.map((f, i)=>{ const hit = map.get(f.featureGroup+"|"+f.featureName); return (
                  <tr key={i} className="border-b last:border-0 hover:bg-slate-50">
                    <td className="p-3 whitespace-nowrap">{f.featureGroup}</td>
                    <td className="p-3">{f.featureName}</td>
                    <td className="p-3">{hit ? badgeForEnabled(hit.enabled) : <Badge className="bg-gray-400">N/A</Badge>}</td>
                    <td className="p-3">{hit?.notes || "—"}</td>
                    <td className="p-3">{!hit && (<>
                      <Button size="sm" variant="outline" onClick={()=> onAddQuick({ featureGroup:f.featureGroup, featureName:f.featureName, patientCategory, modeType, mode, enabled: "Y" })}>Mark Y</Button>
                      <Button size="sm" variant="ghost" onClick={()=> onAddQuick({ featureGroup:f.featureGroup, featureName:f.featureName, patientCategory, modeType, mode, enabled: "N" })}>Mark N</Button>
                    </>)}
                    </td>
                  </tr>
                );})}
              </tbody>
            </table>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function AddFeatureCard({ onAdd }: { onAdd: (r: Row)=>void }) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState<Row>({ featureGroup: "", featureName: "", patientCategory: "", modeType: "", mode: "", enabled: "", notes: "" });
  function submit() { if (!draft.featureName || !draft.patientCategory || !draft.modeType || !draft.mode) return; onAdd({ ...draft, enabled: (draft.enabled || "").toUpperCase(), mode: normalizeMode(draft.mode), modeType: normalizeModeType(draft.modeType) }); setDraft({ featureGroup: "", featureName: "", patientCategory: "", modeType: "", mode: "", enabled: "", notes: "" }); setOpen(false);} 
  return (
    <Card>
      <CardHeader className="pb-2"><CardTitle className="text-base">Add Feature</CardTitle></CardHeader>
      <CardContent className="flex flex-wrap gap-2">
        <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild><Button><Plus className="w-4 h-4 mr-2"/>New Entry</Button></DialogTrigger>
          <DialogContent className="max-w-2xl">
            <DialogHeader><DialogTitle>Add Feature to Combination</DialogTitle></DialogHeader>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 py-2">
              <Input placeholder="Feature Group (e.g., Views)" value={draft.featureGroup} onChange={e=>setDraft({...draft, featureGroup:e.target.value})} />
              <Input placeholder="Feature Name" value={draft.featureName} onChange={e=>setDraft({...draft, featureName:e.target.value})} />
              <Select value={draft.patientCategory} onValueChange={v=>setDraft({...draft, patientCategory: v})}>
                <SelectTrigger><SelectValue placeholder="Patient Category"/></SelectTrigger>
                <SelectContent>{["Adult","Pediatric","Neonatal"].map(c=> <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
              </Select>
              <Select value={draft.modeType} onValueChange={v=>setDraft({...draft, modeType: v})}>
                <SelectTrigger><SelectValue placeholder="Mode Type"/></SelectTrigger>
                <SelectContent>{Object.keys(MODE_CATALOG).map(t=> <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
              </Select>
              <Select value={draft.mode} onValueChange={v=>setDraft({...draft, mode: v})}>
                <SelectTrigger><SelectValue placeholder="Mode"/></SelectTrigger>
                <SelectContent>{(MODE_CATALOG[draft.modeType] || Array.from(new Set(Object.values(MODE_CATALOG).flat()))).map(m => (<SelectItem key={m} value={m}>{m}</SelectItem>))}</SelectContent>
              </Select>
              <Select value={draft.enabled === "" ? NA : draft.enabled} onValueChange={v=>setDraft({...draft, enabled: v === NA ? "" : v})}>
                <SelectTrigger><SelectValue placeholder="Enabled?"/></SelectTrigger>
                <SelectContent><SelectItem value="Y">Enabled</SelectItem><SelectItem value="N">Disabled</SelectItem><SelectItem value={NA}>N/A</SelectItem></SelectContent>
              </Select>
              <div className="md:col-span-2"><Textarea placeholder="Notes" value={draft.notes} onChange={e=>setDraft({...draft, notes:e.target.value})} /></div>
              <div className="md:col-span-2 flex justify-end gap-2 pt-2"><Button variant="outline" onClick={()=>setOpen(false)}>Cancel</Button><Button onClick={submit}>Add</Button></div>
            </div>
          </DialogContent>
        </Dialog>
      </CardContent>
    </Card>
  );
}

function FeatureTable({ rows, editable, onUpdate, onDelete }: { rows: Row[], editable: boolean, onUpdate: (idx:number, patch: Partial<Row>)=>void, onDelete: (idx:number)=>void }) {
  return (
    <Card className="overflow-hidden">
      <CardContent className="p-0">
        <div className="overflow-auto">
          <table className="min-w-full text-sm">
            <thead className="bg-slate-100"><tr><th className="text-left p-3">Feature Group</th><th className="text-left p-3">Feature Name</th><th className="text-left p-3">Patient</th><th className="text-left p-3">Mode Type</th><th className="text-left p-3">Mode</th><th className="text-left p-3">Enabled</th><th className="text-left p-3">Notes</th><th className="text-left p-3">Actions</th></tr></thead>
            <tbody>
              {rows.map((r, idx) => (
                <tr key={idx} className="border-b last:border-0 hover:bg-slate-50">
                  <td className="p-3">{editable ? <InlineInput value={r.featureGroup} onChange={v=>onUpdate(idx,{featureGroup:v})}/> : r.featureGroup}</td>
                  <td className="p-3">{editable ? <InlineInput value={r.featureName} onChange={v=>onUpdate(idx,{featureName:v})}/> : r.featureName}</td>
                  <td className="p-3">{editable ? <InlineInput value={r.patientCategory} onChange={v=>onUpdate(idx,{patientCategory:v})}/> : r.patientCategory}</td>
                  <td className="p-3">{editable ? <InlineInput value={r.modeType} onChange={v=>onUpdate(idx,{modeType:v})}/> : r.modeType}</td>
                  <td className="p-3">{editable ? <InlineInput value={r.mode} onChange={v=>onUpdate(idx,{mode:v})}/> : r.mode}</td>
                  <td className="p-3">{editable ? (
                    <Select value={r.enabled === "" ? NA : r.enabled} onValueChange={v=>onUpdate(idx,{enabled: v === NA ? "" : v})}>
                      <SelectTrigger className="w-[130px]"><SelectValue placeholder="Enabled?"/></SelectTrigger>
                      <SelectContent><SelectItem value="Y">Enabled</SelectItem><SelectItem value="N">Disabled</SelectItem><SelectItem value={NA}>N/A</SelectItem></SelectContent>
                    </Select>
                  ) : badgeForEnabled(r.enabled)}</td>
                  <td className="p-3">{editable ? <InlineTextarea value={r.notes||""} onChange={v=>onUpdate(idx,{notes:v})}/> : (r.notes || "—")}</td>
                  <td className="p-3"><Button size="icon" variant="ghost" onClick={()=>onDelete(idx)} title="Delete"><Trash2 className="w-4 h-4"/></Button></td>
                </tr>
              ))}
              {!rows.length && (<tr><td colSpan={8} className="p-6 text-center text-slate-500">No rows match your filters.</td></tr>)}
            </tbody>
          </table>
        </div>
      </CardContent>
    </Card>
  );
}

function InlineInput({ value, onChange }: { value: string, onChange: (v:string)=>void }) { return <Input value={value} onChange={e=>onChange(e.target.value)} className="h-8"/>; }
function InlineTextarea({ value, onChange }: { value: string, onChange: (v:string)=>void }) { return <Textarea value={value} onChange={e=>onChange(e.target.value)} className="h-16"/>; }

function FeatureMatrix({ rows }: { rows: Row[] }) {
  const features = Array.from(new Map(rows.map(r => [r.featureGroup+"|"+r.featureName, { group: r.featureGroup, name: r.featureName }])).values());
  const combos = Array.from(new Map(rows.map(r => [r.patientCategory+"|"+r.modeType+"|"+r.mode, { patient: r.patientCategory, type: r.modeType, mode: r.mode }])).values());
  return (
    <div className="overflow-auto">
      <table className="min-w-full text-sm">
        <thead><tr><th className="p-3 bg-slate-100 sticky left-0 z-10">Feature</th>{combos.map((c,i)=>(<th key={i} className="p-3 bg-slate-100 whitespace-nowrap">{c.patient} / {c.type} / {c.mode}</th>))}</tr></thead>
        <tbody>
          {features.map((f, fi)=> (
            <tr key={fi} className="border-b last:border-0">
              <td className="p-3 bg-white sticky left-0 z-10"><div className="font-medium">{f.name}</div><div className="text-xs text-slate-500">{f.group}</div></td>
              {combos.map((c, ci)=>{ const hit = rows.find(r => r.featureGroup===f.group && r.featureName===f.name && r.patientCategory===c.patient && r.modeType===c.type && r.mode===c.mode); return (<td key={ci} className="p-2 text-center">{hit ? badgeForEnabled(hit.enabled) : <Badge className="bg-gray-400">—</Badge>}</td>); })}
            </tr>
          ))}
          {!features.length && (<tr><td className="p-6 text-center text-slate-500" colSpan={combos.length+1}>No data for the selected filters.</td></tr>)}
        </tbody>
      </table>
    </div>
  );
}

function SelfTests({ rows }: { rows: Row[] }) {
  const [results, setResults] = useState<string[]>([]);
  function run() {
    const out: string[] = [];
    out.push((ALL !== "" && NA !== "") ? "PASS: Sentinels are non-empty" : "FAIL: Sentinel values must not be empty");
    const cases: [string,string][] = [["pcv","PC"],["vcv","VC"],["simv pc","SIMV+PC"],["SIMV-VC","SIMV+VC"],["SIMV PRVC","SIMV+PRVC"],["CPAP + PS","PS/CPAP"],["vs vc","VS-VC"],["vs prvc","VS-PRVC"],["HFNC","HFOT"],["niv ps","NIV-PS"],["NCPAP","nCPAP"],["Dual Pap","DualPap"]];
    out.push(cases.every(([inp,exp]) => normalizeMode(inp) === exp) ? "PASS: normalizeMode canonicalizes expected aliases" : "FAIL: normalizeMode mapping issue");
    const f = (m: string, c: string, t: string) => rows.filter(r => (!m || normalizeMode(r.mode)===m) && (!c || r.patientCategory===c) && (!t || normalizeModeType(r.modeType)===t));
    const before = f("", "", "").length; const after = f("PC", "Adult", "Invasive").length; out.push(before >= after ? "PASS: Filter reduces or equals dataset size" : "FAIL: Filter increased dataset unexpectedly");
    const want = new Set(["PC","VC","PRVC","SIMV+PC","SIMV+VC","SIMV+PRVC","PS/CPAP","VS-VC","VS-PRVC","APRV","NIV-PS","HFOT","DualPap","NIV-PC","nCPAP"]);
    const have = new Set(Object.values(MODE_CATALOG).flat()); const missing = Array.from(want).filter(w=>!have.has(w)); const extras  = Array.from(have).filter(h=>!want.has(h));
    out.push(missing.length===0 && extras.length===0 ? "PASS: Catalog exactly matches provided modes" : `FAIL: Catalog mismatch. Missing: ${missing.join(", ")} | Extra: ${extras.join(", ")}`);
    setResults(out);
  }
  return (
    <Card className="mt-6"><CardHeader className="pb-2"><CardTitle className="text-base">Dev: Self-tests</CardTitle></CardHeader><CardContent>
      <p className="text-sm text-slate-600 mb-3">Checks for dropdown sentinels, alias normalization, catalog coverage, and filtering.</p>
      <div className="flex gap-2 mb-3"><Button variant="outline" onClick={run}>Run Tests</Button></div>
      <ul className="list-disc pl-6 text-sm">{results.map((r,i)=>(<li key={i}>{r}</li>))}{!results.length && <li className="text-slate-500">(No results yet)</li>}</ul>
    </CardContent></Card>
  );
}
