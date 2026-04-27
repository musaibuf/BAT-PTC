import React, { useState, useEffect, useCallback, useRef } from "react";
import pptxgen from "pptxgenjs";
import {
  ThemeProvider, createTheme, CssBaseline,
  Box, Typography, Button, TextField, Paper, Grid,
  Chip, LinearProgress, Divider, AppBar, Toolbar,
  Alert, CircularProgress, Stack, Fade, Accordion, AccordionSummary, AccordionDetails
} from "@mui/material";

import ArrowBackIcon from "@mui/icons-material/ArrowBack";
import NavigateNextIcon from "@mui/icons-material/NavigateNext";
import NavigateBeforeIcon from "@mui/icons-material/NavigateBefore";
import AutoFixHighIcon from "@mui/icons-material/AutoFixHigh";
import MarkEmailReadIcon from "@mui/icons-material/MarkEmailRead";
import PrintIcon from "@mui/icons-material/Print";
import GroupsIcon from "@mui/icons-material/Groups";
import AdminPanelSettingsIcon from "@mui/icons-material/AdminPanelSettings";
import CheckCircleOutlinedIcon from "@mui/icons-material/CheckCircleOutlined";
import WarningAmberIcon from "@mui/icons-material/WarningAmber";
import RefreshIcon from "@mui/icons-material/Refresh";
import AssessmentIcon from "@mui/icons-material/Assessment";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import DownloadIcon from "@mui/icons-material/Download";

// ─── THEME ────────────────────────────────────────────────────────────────────
const BAT_NAVY  = "#17468B";
const BAT_GOLD  = "#FAB41E";
const BAT_DARK  = "#0d2a5a";
const BAT_LIGHT = "#EEF3FB";
const CARN_RED  = "#C0392B";

const theme = createTheme({
  palette: {
    mode: "light",
    primary:   { main: BAT_NAVY, dark: BAT_DARK, light: "#2660C0", contrastText: "#fff" },
    secondary: { main: BAT_GOLD, dark: "#d4940a", contrastText: "#0d2a5a" },
    error:     { main: CARN_RED },
    background:{ default: "#F4F6FB", paper: "#fff" },
    text:      { primary: "#0d2a5a", secondary: "#4A6080" },
  },
  typography: {
    fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif",
    h1: { fontWeight: 800, letterSpacing: "-0.03em" },
    h2: { fontWeight: 700, letterSpacing: "-0.02em" },
    h3: { fontWeight: 700 },
    h4: { fontWeight: 700 },
    h5: { fontWeight: 600 },
    h6: { fontWeight: 600 },
    overline: { fontWeight: 700, letterSpacing: "0.15em", fontSize: "0.65rem" },
    caption: { letterSpacing: "0.05em" },
  },
  shape: { borderRadius: 4 },
  components: {
    MuiButton: {
      styleOverrides: {
        root: { textTransform: "none", fontWeight: 700, letterSpacing: "0.02em", borderRadius: 4 },
        containedPrimary: { background: `linear-gradient(135deg, ${BAT_NAVY} 0%, ${BAT_DARK} 100%)`, "&:hover": { background: `linear-gradient(135deg, #2660C0 0%, ${BAT_NAVY} 100%)` }, boxShadow: `0 2px 12px rgba(23,70,139,0.35)` },
        containedSecondary: { background: `linear-gradient(135deg, ${BAT_GOLD} 0%, #e09900 100%)`, "&:hover": { background: `linear-gradient(135deg, #ffc84a 0%, ${BAT_GOLD} 100%)` }, boxShadow: `0 2px 10px rgba(250,180,30,0.4)` },
      },
    },
    MuiPaper: { styleOverrides: { root: { backgroundImage: "none" }, elevation1: { boxShadow: "0 1px 4px rgba(13,42,90,0.08), 0 4px 16px rgba(13,42,90,0.06)" }, elevation3: { boxShadow: "0 4px 20px rgba(13,42,90,0.12), 0 1px 4px rgba(13,42,90,0.08)" } } },
    MuiTextField: { styleOverrides: { root: { "& .MuiOutlinedInput-root": { "&:hover fieldset": { borderColor: BAT_NAVY }, "&.Mui-focused fieldset": { borderColor: BAT_NAVY, borderWidth: 2 } }, "& label.Mui-focused": { color: BAT_NAVY } } } },
    MuiLinearProgress: { styleOverrides: { root: { borderRadius: 2, height: 5 } } },
    MuiChip: { styleOverrides: { root: { fontWeight: 600, fontSize: "0.7rem" } } },
    MuiDivider: { styleOverrides: { root: { borderColor: "rgba(23,70,139,0.12)" } } },
  },
});

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const API_URL = process.env.REACT_APP_API_URL || "http://localhost:5000/api";

const OPM = [
  { key: "task",         label: "Task",          def: "Are specific tasks clearly identified which lead to achieving the strategy and goals?", prompt: "Today — are priorities clear? What gets called urgent vs what actually matters?", futurePrompt: "Based on your 2027 headline, what specific tasks and priorities will be clearly identified to achieve this new reality?" },
  { key: "people",       label: "People",        def: "Do people have the skills to do the tasks? Training, hiring, promotion.", prompt: "Today — do people have the skills needed? What's happening with hiring, training, promotions?", futurePrompt: "To make this 2027 reality work, what new skills will your people have? How will hiring and training have changed?" },
  { key: "structure",    label: "Structure",     def: "Does the structure let the right people work together?", prompt: "Today — do the right people work together easily? Where do silos break down?", futurePrompt: "How will your structure have shifted by 2027 to allow the right people to work together and break down today's silos?" },
  { key: "decisionMaking", label: "Decision Making", def: "Do decisions reflect knowledge and experience? Level of delegation.", prompt: "Today — who makes which decisions? Where does delegation work, where does it stall?", futurePrompt: "In this 2027 future, who will make which decisions? How will delegation have improved?" },
  { key: "information",  label: "Information",   def: "Is needed information available? Valid feedback, timely data.", prompt: "Today — what do people know, what do they find out too late?", futurePrompt: "What information and data will flow freely in 2027 that is currently missing or delayed today?" },
  { key: "rewards",      label: "Rewards",       def: "Are desired behaviors rewarded? Undesired behaviors punished?", prompt: "Today — what behaviors actually get rewarded here, whatever the policy says?", futurePrompt: "By 2027, what specific behaviors will actually be rewarded to sustain this new culture?" },
];

const PHASES      = ["setup", "phase1", "phase2", "phase3", "phase4", "complete"];
const PHASE_LABEL = { setup: "Awaiting Start", phase1: "Current State", phase2: "Newspaper 2027", phase3: "Reverse-Engineer", phase4: "Roadblocks", complete: "Synthesis" };

// ─── API ──────────────────────────────────────────────────────────────────────
async function loadSession(code) { if (!code) return null; try { const r = await fetch(`${API_URL}/sessions/${code}`); if (!r.ok) return null; return await r.json(); } catch { return null; } }
async function createSession(data) { try { await fetch(`${API_URL}/sessions`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(data) }); return true; } catch { return false; } }
async function updateSessionPhase(code, phase) { try { await fetch(`${API_URL}/sessions/${code}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify({currentPhase:phase}) }); } catch(e) { console.error(e); } }
async function loadAllGroups(code) { try { const r = await fetch(`${API_URL}/sessions/${code}/groups`); return await r.json(); } catch { return {}; } }
async function loadGroup(code, num) { try { const r = await fetch(`${API_URL}/sessions/${code}/groups/${num}`); if (!r.ok) return null; return await r.json(); } catch { return null; } }
async function saveGroup(code, num, data) { try { await fetch(`${API_URL}/sessions/${code}/groups/${num}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(data) }); return true; } catch { return false; } }
async function endSession(code) { try { await fetch(`${API_URL}/sessions/${code}`, { method:"DELETE" }); } catch(e) { console.error(e); } }

async function checkNudge(label, answer) {
  if (!answer || answer.trim().split(/\s+/).length < 3) return { needsNudge:false, probe:"" };
  const sys = `You check if a senior leader's answer in a culture workshop is specific enough. Return ONLY valid JSON: {"needsNudge":true/false,"probe":"..."}. Set needsNudge=true ONLY when: Answer is under 12 words AND not self-evidently complete, OR Uses corporate abstractions without specifics. If needsNudge=true, probe is ONE sharp follow-up under 18 words.`;
  try {
    const r = await fetch(`${API_URL}/ai/nudge`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ system:sys, user:`Question: "${label}"\nAnswer: "${answer}"` }) });
    const d = await r.json(); const p = JSON.parse(d.text.replace(/```json|```/g,"").trim());
    return { needsNudge:!!p.needsNudge, probe:p.probe||"" };
  } catch { return { needsNudge:false, probe:"" }; }
}

async function autoMap(np) {
  const sys = `Map a 2027 future-state newspaper article to 6 organizational design elements. Return ONLY valid JSON with keys: task, people, structure, decisionMaking, information, rewards. Each value = ONE concrete sentence. If not addressed, write exactly "Not addressed in this article."`;
  const user = `HEADLINE: ${np.headline}\nACTION 1: ${np.action1}\nACTION 2: ${np.action2}\nACTION 3: ${np.action3}\nFRONTLINE QUOTE: "${np.frontlineQuote}"`;
  const out = {}; for (const e of OPM) out[e.key]="";
  try {
    const r = await fetch(`${API_URL}/ai/automap`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({system:sys,user}) });
    const d = await r.json(); const p = JSON.parse(d.text.replace(/```json|```/g,"").trim());
    for (const e of OPM) out[e.key]=p[e.key]||"Not addressed in this article.";
    return out;
  } catch { return out; }
}

async function generateSummary(data) {
  try {
    const r = await fetch(`${API_URL}/ai/summarize`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({data: JSON.stringify(data)}) });
    const d = await r.json(); return d.text;
  } catch { return "Summary unavailable."; }
}

async function generateThemes(data) {
  try {
    const r = await fetch(`${API_URL}/ai/themes`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({data: JSON.stringify(data)}) });
    const d = await r.json(); return JSON.parse(d.text.replace(/```json|```/g,"").trim());
  } catch { return { executiveSummary: "Summary unavailable.", themes: ["Theme 1", "Theme 2", "Theme 3"] }; }
}

function genCode() { const c="ABCDEFGHJKLMNPQRSTUVWXYZ"; let s=""; for (let i=0;i<4;i++) s+=c[Math.floor(Math.random()*c.length)]; return s; }

// ─── HEADER ───────────────────────────────────────────────────────────────────
function Header({ code, group, phase }) {
  return (
    <AppBar position="static" elevation={0} sx={{ background:`linear-gradient(135deg, ${BAT_DARK} 0%, ${BAT_NAVY} 60%, #1e56a8 100%)`, borderBottom:`3px solid ${BAT_GOLD}` }}>
      <Toolbar sx={{ minHeight:64, px:{ xs:2, md:4 } }}>
        <Box component="img" src="/BAT.png" alt="BAT" sx={{ height:38, mr:2, filter:"brightness(0) invert(1)" }} onError={e=>{ e.target.style.display="none"; }} />
        <Box sx={{ flex:1 }}>
          <Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", lineHeight:1, mb:0.2 }}>Culture Gap Assessment</Typography>
          <Typography variant="h6" sx={{ color:"#fff", fontWeight:800, lineHeight:1, fontSize:"1rem", letterSpacing:"-0.01em" }}>Organisational Performance Model</Typography>
        </Box>
        <Stack direction="row" spacing={1.5} alignItems="center">
          {code && <Box sx={{ textAlign:"center", px:1.5, py:0.5, border:`1px solid rgba(250,180,30,0.4)`, borderRadius:1 }}><Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", lineHeight:1, fontSize:"0.6rem" }}>Session</Typography><Typography variant="h6" sx={{ color:"#fff", fontWeight:900, letterSpacing:"0.2em", lineHeight:1 }}>{code}</Typography></Box>}
          {group && <Chip label={`Group ${group}`} size="small" sx={{ background:BAT_GOLD, color:BAT_DARK, fontWeight:800, fontSize:"0.75rem" }} />}
          {phase && <Chip label={PHASE_LABEL[phase]||phase} size="small" variant="outlined" sx={{ borderColor:"rgba(255,255,255,0.3)", color:"rgba(255,255,255,0.85)", fontSize:"0.65rem" }} />}
          <Box sx={{ display:"flex", alignItems:"center", gap:0.75, pl:1.5, borderLeft:"1px solid rgba(255,255,255,0.2)" }}>
            <Box component="img" src="/logo.png" alt="Carnelian Co" sx={{ height:22, opacity:0.85 }} onError={e=>{ e.target.style.display="none"; }} />
            <Typography variant="caption" sx={{ color:"rgba(255,255,255,0.55)", fontSize:"0.6rem", lineHeight:1 }}>Powered by<br /><b style={{color:"rgba(255,255,255,0.8)"}}>Carnelian Co</b></Typography>
          </Box>
        </Stack>
      </Toolbar>
    </AppBar>
  );
}

// ─── PHASE STEPPER ────────────────────────────────────────────────────────────
function PhaseStepper({ current }) {
  const steps = ["Current State","Newspaper","Mapping","Roadblocks","Synthesis"];
  const keys  = ["phase1","phase2","phase3","phase4","complete"];
  const idx   = keys.indexOf(current);
  return (
    <Box sx={{ display:"flex", mb:3, border:`1px solid rgba(23,70,139,0.15)`, borderRadius:1, overflow:"hidden" }}>
      {steps.map((s,i)=>{
        const done = i < idx, active = i === idx;
        return (
          <Box key={s} sx={{ flex:1, py:1, px:0.5, textAlign:"center", background: active ? BAT_NAVY : done ? BAT_LIGHT : "#fff", borderRight: i<steps.length-1 ? `1px solid rgba(23,70,139,0.12)` : "none", transition:"background 0.3s" }}>
            <Typography variant="overline" sx={{ color: active ? BAT_GOLD : done ? BAT_NAVY : "#9bb0cc", fontSize:"0.6rem", display:"block", lineHeight:1.2 }}>{i+1} · {s}</Typography>
          </Box>
        );
      })}
    </Box>
  );
}

// ─── SMART FIELD ──────────────────────────────────────────────────────────────
function SmartField({ label, value, onChange, placeholder, rows=3, ctx }) {
  const [probe, setProbe] = useState("");
  const [checking, setChecking] = useState(false);
  const [dismissed, setDismissed] = useState(false);

  const check = async () => {
    if (!value || value.trim().split(/\s+/).length < 3 || dismissed) return;
    setChecking(true); const r = await checkNudge(ctx||label, value); setChecking(false);
    if (r.needsNudge && r.probe) setProbe(r.probe);
  };

  return (
    <Box sx={{ mb:2 }}>
      <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5 }}>{label}</Typography>
      <TextField fullWidth multiline rows={rows} value={value} onChange={e=>onChange(e.target.value)} onBlur={check} placeholder={placeholder||"Be specific…"} variant="outlined" size="small" sx={{ "& .MuiOutlinedInput-root":{ fontSize:"0.9rem" } }} />
      {checking && <Typography variant="caption" sx={{ color:BAT_NAVY, display:"block", mt:0.5 }}><CircularProgress size={10} sx={{ mr:0.5, color:BAT_NAVY }} />Checking depth…</Typography>}
      {probe && !dismissed && (
        <Fade in>
          <Alert severity="warning" onClose={()=>{setDismissed(true);setProbe("");}} icon={<AutoFixHighIcon fontSize="small"/>} sx={{ mt:1, borderLeft:`3px solid ${BAT_GOLD}`, background:"#fffbea", "& .MuiAlert-message":{ fontSize:"0.82rem", fontStyle:"italic", color:"#5a4000" } }}>{probe}</Alert>
        </Fade>
      )}
    </Box>
  );
}

// ─── LANDING ──────────────────────────────────────────────────────────────────
function Landing({ go }) {
  return (
    <Box>
      <Header />
      <Box sx={{ maxWidth:760, mx:"auto", px:2, pt:6, pb:10 }}>
        <Box sx={{ textAlign:"center", mb:6 }}>
          <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>BAT × Carnelian · Workshop Tool</Typography>
          <Typography variant="h3" sx={{ color:BAT_DARK, fontWeight:800, mt:1, mb:2, lineHeight:1.15 }}>Culture Gap Assessment</Typography>
          <Typography sx={{ color:"#4A6080", maxWidth:500, mx:"auto", lineHeight:1.7 }}>A structured workshop to surface the gap between today's organisational reality and the 2027 ambition — across six design elements.</Typography>
        </Box>
        <Grid container spacing={2.5}>
          <Grid item xs={12} sm={6}>
            <Paper elevation={3} sx={{ p:3.5, borderTop:`3px solid ${BAT_NAVY}`, height:"100%", display:"flex", flexDirection:"column" }}>
              <Box sx={{ display:"flex", alignItems:"center", gap:1.5, mb:2 }}>
                <Box sx={{ width:42, height:42, borderRadius:1, background:`linear-gradient(135deg,${BAT_NAVY},${BAT_DARK})`, display:"flex", alignItems:"center", justifyContent:"center" }}><AdminPanelSettingsIcon sx={{ color:"#fff", fontSize:22 }} /></Box>
                <Box><Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", lineHeight:1 }}>Role</Typography><Typography variant="h6" sx={{ color:BAT_DARK, fontWeight:800 }}>Facilitator</Typography></Box>
              </Box>
              <Typography variant="body2" sx={{ color:"#4A6080", mb:3, flex:1, lineHeight:1.7 }}>Set up the session, share the access code with the room, and advance groups through each phase from your dashboard.</Typography>
              <Button fullWidth variant="contained" color="primary" size="large" endIcon={<NavigateNextIcon/>} onClick={()=>go("fac")}>Start as Facilitator</Button>
            </Paper>
          </Grid>
          <Grid item xs={12} sm={6}>
            <Paper elevation={3} sx={{ p:3.5, borderTop:`3px solid ${BAT_GOLD}`, height:"100%", display:"flex", flexDirection:"column" }}>
              <Box sx={{ display:"flex", alignItems:"center", gap:1.5, mb:2 }}>
                <Box sx={{ width:42, height:42, borderRadius:1, background:`linear-gradient(135deg,${BAT_GOLD},#d4940a)`, display:"flex", alignItems:"center", justifyContent:"center" }}><GroupsIcon sx={{ color:BAT_DARK, fontSize:22 }} /></Box>
                <Box><Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", lineHeight:1 }}>Role</Typography><Typography variant="h6" sx={{ color:BAT_DARK, fontWeight:800 }}>Group Scribe</Typography></Box>
              </Box>
              <Typography variant="body2" sx={{ color:"#4A6080", mb:3, flex:1, lineHeight:1.7 }}>One device per breakout table. Enter the session code and group number to capture your group's answers in real time.</Typography>
              <Button fullWidth variant="contained" color="secondary" size="large" endIcon={<NavigateNextIcon/>} onClick={()=>go("grp")}>Join as Group Scribe</Button>
            </Paper>
          </Grid>
        </Grid>
        <Box sx={{ mt:6, pt:3, borderTop:`1px solid rgba(23,70,139,0.12)`, display:"flex", alignItems:"center", justifyContent:"center", gap:2 }}>
          <Box component="img" src="/BAT.png" alt="BAT" sx={{ height:28, opacity:0.7 }} onError={e=>{e.target.style.display="none";}} />
          <Divider orientation="vertical" flexItem />
          <Box component="img" src="/logo.png" alt="Carnelian Co" sx={{ height:22, opacity:0.6, filter: "brightness(0)" }} onError={e=>{e.target.style.display="none";}} />
          <Typography variant="caption" sx={{ color:"#9bb0cc" }}>© Carnelian · carnelianco.com</Typography>
        </Box>
      </Box>
    </Box>
  );
}

// ─── FACILITATOR ──────────────────────────────────────────────────────────────
function Facilitator({ onExit }) {
  const [sessionCode, setSessionCode] = useState(localStorage.getItem("fac_session")||null);
  const [session, setSession]         = useState(null);
  const [groups, setGroups]           = useState({});
  const [loading, setLoading]         = useState(true);
  const [report, setReport]           = useState(false);

  const refresh = useCallback(async ()=>{
    if (!sessionCode) { setLoading(false); return; }
    const s = await loadSession(sessionCode);
    if (s) { setSession(s); const g = await loadAllGroups(sessionCode); setGroups(g); }
    else { setSessionCode(null); localStorage.removeItem("fac_session"); }
    setLoading(false);
  },[sessionCode]);

  useEffect(()=>{ refresh(); const t=setInterval(refresh,8000); return()=>clearInterval(t); },[refresh]);

  const handleCreated=(code)=>{ localStorage.setItem("fac_session",code); setSessionCode(code); setLoading(true); refresh(); };
  const handleExit=()=>{ localStorage.removeItem("fac_session"); setSessionCode(null); onExit(); };

  if (loading) return <Box sx={{display:"flex",justifyContent:"center",pt:10}}><CircularProgress sx={{color:BAT_NAVY}}/></Box>;
  if (!sessionCode||!session) return <FacSetup onCreated={handleCreated} onExit={handleExit}/>;
  if (report) return <Report session={session} groups={groups} onClose={()=>setReport(false)}/>;
  return <FacDash session={session} groups={groups} refresh={refresh} showReport={()=>setReport(true)} onExit={handleExit}/>;
}

function FacSetup({ onCreated, onExit }) {
  const [num,setNum]   = useState(4);
  const [name,setName] = useState("BAT — Culture Gap Assessment");
  const [busy,setBusy] = useState(false);

  const create = async()=>{
    setBusy(true); const code=genCode();
    await createSession({ code, numGroups:num, sessionName:name, currentPhase:"setup", startedAt:Date.now() });
    onCreated(code); setBusy(false);
  };

  return (
    <Box>
      <Header />
      <Box sx={{ maxWidth:560, mx:"auto", px:2, pt:5 }}>
        <Button startIcon={<ArrowBackIcon/>} onClick={onExit} sx={{ mb:3, color:BAT_NAVY }}>Back</Button>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Facilitator Setup</Typography>
        <Typography variant="h4" sx={{ color:BAT_DARK, fontWeight:800, mb:3, mt:0.5 }}>Create a new session</Typography>
        <Paper elevation={1} sx={{ p:3.5 }}>
          <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5 }}>Session Name</Typography>
          <TextField fullWidth value={name} onChange={e=>setName(e.target.value)} variant="outlined" sx={{ mb:3, "& input":{ fontWeight:700, fontSize:"1.05rem" } }} />
          <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:1 }}>Number of Groups</Typography>
          <Stack direction="row" spacing={1} sx={{ mb:3.5 }}>
            {[3,4,5,6].map(n=><Button key={n} variant={num===n?"contained":"outlined"} color="primary" onClick={()=>setNum(n)} sx={{ flex:1, fontWeight:800, fontSize:"1rem" }}>{n}</Button>)}
          </Stack>
          <Button fullWidth variant="contained" color="primary" size="large" disabled={busy||!name.trim()} onClick={create} endIcon={busy ? <CircularProgress size={16} sx={{color:"#fff"}}/> : <NavigateNextIcon/>} sx={{ py:1.5 }}>Create Session</Button>
        </Paper>
      </Box>
    </Box>
  );
}

function FacDash({ session, groups, refresh, showReport, onExit }) {
  const advance = async()=>{ const i=PHASES.indexOf(session.currentPhase); if(i>=PHASES.length-1)return; await updateSessionPhase(session.code,PHASES[i+1]); await refresh(); };
  const back    = async()=>{ const i=PHASES.indexOf(session.currentPhase); if(i<=0)return; await updateSessionPhase(session.code,PHASES[i-1]); await refresh(); };
  const end     = async()=>{ if(!window.confirm("End session? All data will be cleared."))return; await endSession(session.code); onExit(); };

  const pNum     = PHASES.indexOf(session.currentPhase);
  const pLabel   = PHASE_LABEL[session.currentPhase];
  const nextLabel= PHASE_LABEL[PHASES[pNum+1]];

  return (
    <Box>
      <Header code={session.code} phase={session.currentPhase} />
      <Box sx={{ background:`linear-gradient(90deg,${BAT_DARK}f5,${BAT_NAVY}f5)`, px:{ xs:2, md:4 }, py:1.5, display:"flex", alignItems:"center", gap:2, flexWrap:"wrap", borderBottom:`2px solid ${BAT_GOLD}44` }}>
        <Typography sx={{ color:"#fff", fontWeight:700, fontSize:"0.85rem", flex:1 }}>Phase {pNum} · {pLabel}</Typography>
        <Stack direction="row" spacing={1}>
          <Button size="small" variant="outlined" startIcon={<NavigateBeforeIcon/>} onClick={back} sx={{ color:"#fff", borderColor:"rgba(255,255,255,0.3)", "&:hover":{borderColor:"#fff",background:"rgba(255,255,255,0.1)"} }}>Prev</Button>
          <Button size="small" variant="contained" color="secondary" endIcon={<NavigateNextIcon/>} onClick={advance} disabled={session.currentPhase==="complete"}>{nextLabel ? `→ ${nextLabel}` : "Advance"}</Button>
          <Button size="small" onClick={refresh} sx={{ color:"#fff", minWidth:36 }}><RefreshIcon fontSize="small"/></Button>
        </Stack>
        <Divider orientation="vertical" flexItem sx={{ borderColor:"rgba(255,255,255,0.2)" }}/>
        <Button size="small" variant="outlined" startIcon={<AssessmentIcon/>} onClick={showReport} sx={{ color:BAT_GOLD, borderColor:BAT_GOLD, "&:hover":{background:"rgba(250,180,30,0.1)"} }}>Report</Button>
        <Button size="small" onClick={end} sx={{ color:"rgba(255,255,255,0.5)", fontSize:"0.75rem" }}>End Session</Button>
      </Box>

      <Box sx={{ maxWidth:1200, mx:"auto", px:{ xs:2, md:4 }, pt:3, pb:8 }}>
        <Box sx={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", mb:3, flexWrap:"wrap", gap:2 }}>
          <Box><Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Live Session</Typography><Typography variant="h5" sx={{ color:BAT_DARK, fontWeight:800, mt:0.2 }}>{session.sessionName}</Typography></Box>
          <Paper elevation={2} sx={{ px:3, py:1.5, textAlign:"center", borderTop:`3px solid ${BAT_GOLD}` }}><Typography variant="overline" sx={{ color:"#4A6080", display:"block", lineHeight:1 }}>Room Code</Typography><Typography sx={{ fontWeight:900, fontSize:"2rem", letterSpacing:"0.25em", color:BAT_DARK, lineHeight:1 }}>{session.code}</Typography></Paper>
        </Box>

        <Typography variant="overline" sx={{ color:BAT_NAVY, fontWeight:700, display:"block", mb:1.5 }}>Group Progress · {session.numGroups} groups</Typography>
        <Grid container spacing={2} sx={{ mb:4 }}>
          {Array.from({length:session.numGroups},(_,i)=>i+1).map(n=>{
            const g = groups[n]; const joined = g?.joined;
            const cnt=(ph)=>g?.[ph]?Object.values(g[ph]).filter(v=>v&&String(v).trim()).length:0;
            return (
              <Grid item xs={12} sm={6} md={4} lg={3} key={n}>
                <Paper elevation={joined?3:1} sx={{ p:2.5, height:"100%", borderTop:`3px solid ${joined?BAT_GOLD:BAT_LIGHT}`, outline: joined?`1px solid ${BAT_GOLD}44`:"none" }}>
                  <Box sx={{ display:"flex", alignItems:"center", justifyContent:"space-between", mb:2 }}>
                    <Typography variant="h6" sx={{ fontWeight:800, color:BAT_DARK }}>Group {n}</Typography>
                    <Chip label={joined?"Live":"Waiting"} size="small" sx={{ background:joined?`${BAT_GOLD}22`:BAT_LIGHT, color:joined?BAT_DARK:"#9bb0cc", fontWeight:700, border:`1px solid ${joined?BAT_GOLD:"#dde5f0"}` }} />
                  </Box>
                  <Stack spacing={0.8} sx={{ mb:2 }}>
                    <ProgressRow label="Current" done={cnt("phase1")} total={6}/>
                    <ProgressRow label="Newspaper" done={cnt("phase2")} total={5}/>
                    <ProgressRow label="Mapping" done={cnt("phase3")} total={6} reviewed={g?.phase3Reviewed}/>
                    <ProgressRow label="Roadblocks" done={cnt("phase4")} total={3}/>
                  </Stack>
                  {/* Live Summaries Accordion */}
                  {joined && <GroupSummaries sessionCode={session.code} group={g} refresh={refresh} />}
                </Paper>
              </Grid>
            );
          })}
        </Grid>
        {(session.currentPhase==="phase3"||session.currentPhase==="phase4"||session.currentPhase==="complete") && <Synth session={session} groups={groups}/>}
      </Box>
    </Box>
  );
}

function ProgressRow({ label, done, total, reviewed }) {
  const complete = done >= total;
  return (
    <Box>
      <Box sx={{ display:"flex", justifyContent:"space-between", mb:0.3 }}>
        <Typography variant="caption" sx={{ color:complete?"#166534":"#4A6080", fontWeight:600 }}>{label}</Typography>
        <Typography variant="caption" sx={{ color:complete?"#166534":"#9bb0cc", fontFamily:"monospace" }}>{reviewed&&complete?"✓ reviewed":`${done}/${total}`}</Typography>
      </Box>
      <LinearProgress variant="determinate" value={(done/total)*100} sx={{ "& .MuiLinearProgress-bar":{ background:complete?`linear-gradient(90deg,#166534,#27a060)`:`linear-gradient(90deg,${BAT_NAVY},#2660C0)` }, background: complete?"#dcf5e7":"#EEF3FB" }}/>
    </Box>
  );
}

function GroupSummaries({ sessionCode, group, refresh }) {
  const [loading, setLoading] = useState(false);

  const handleSummarize = async (phaseKey) => {
    setLoading(true);
    const text = await generateSummary(group[phaseKey]);
    const newSummaries = { ...(group.summaries || {}), [phaseKey]: text };
    await saveGroup(sessionCode, group.groupNumber, { summaries: newSummaries });
    await refresh();
    setLoading(false);
  };

  const phases = [
    { key: "phase1", label: "Phase 1: Today" },
    { key: "phase2", label: "Phase 2: Newspaper" },
    { key: "phase3", label: "Phase 3: Mapping" },
    { key: "phase4", label: "Phase 4: Roadblocks" }
  ];

  return (
    <Accordion elevation={0} sx={{ background: BAT_LIGHT, border: `1px solid rgba(23,70,139,0.1)`, "&:before": { display: "none" } }}>
      <AccordionSummary expandIcon={<ExpandMoreIcon />} sx={{ minHeight: 36, "& .MuiAccordionSummary-content": { my: 1 } }}>
        <Typography variant="caption" sx={{ fontWeight: 700, color: BAT_NAVY }}>View AI Summaries</Typography>
      </AccordionSummary>
      <AccordionDetails sx={{ pt: 0, px: 2, pb: 2 }}>
        <Stack spacing={1.5}>
          {phases.map(p => {
            const hasData = group[p.key] && Object.values(group[p.key]).filter(v => v && String(v).trim()).length > 0;
            const summary = group.summaries?.[p.key];
            if (!hasData) return null;
            return (
              <Box key={p.key} sx={{ background: "#fff", p: 1.5, borderRadius: 1, border: "1px solid #e0e6ed" }}>
                <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
                  <Typography variant="overline" sx={{ color: BAT_DARK, lineHeight: 1 }}>{p.label}</Typography>
                  {!summary && <Button size="small" variant="text" onClick={() => handleSummarize(p.key)} disabled={loading} sx={{ minWidth: 0, p: 0, fontSize: "0.65rem" }}>{loading ? "..." : "Summarize"}</Button>}
                </Box>
                {summary ? <Typography variant="body2" sx={{ color: "#4A6080", fontSize: "0.75rem", lineHeight: 1.4 }}>{summary}</Typography> : <Typography variant="body2" sx={{ color: "#9bb0cc", fontSize: "0.7rem", fontStyle: "italic" }}>No summary yet.</Typography>}
              </Box>
            );
          })}
        </Stack>
      </AccordionDetails>
    </Accordion>
  );
}

function Synth({ session, groups }) {
  const data = OPM.map(el=>{
    const current=[],future=[];
    for(let i=1;i<=session.numGroups;i++){
      const g=groups[i];
      if(g?.phase1?.[el.key]) current.push({g:i,t:g.phase1[el.key]});
      if(g?.phase3?.[el.key]) future.push({g:i,t:g.phase3[el.key]});
    }
    return {...el,current,future};
  });

  return (
    <Box>
      <Typography variant="overline" sx={{ color:BAT_NAVY, fontWeight:700, display:"block", mb:0.5 }}>Live Synthesis</Typography>
      <Typography variant="h5" sx={{ color:BAT_DARK, fontWeight:800, mb:2.5 }}>The Gap Across Six Design Elements</Typography>
      <Stack spacing={1.5}>
        {data.map(el=>(
          <Paper key={el.key} elevation={1} sx={{ overflow:"hidden" }}>
            <Box sx={{ px:2.5, py:1.5, background:`linear-gradient(90deg,${BAT_LIGHT},#fff)`, borderBottom:`1px solid rgba(23,70,139,0.08)`, display:"flex", alignItems:"center", gap:2 }}>
              <Typography variant="h6" sx={{ fontWeight:800, color:BAT_DARK, flex:1 }}>{el.label}</Typography>
              <Typography variant="caption" sx={{ fontFamily:"monospace", color:"#4A6080" }}>{el.current.length}C / {el.future.length}F</Typography>
            </Box>
            <Grid container>
              <Grid item xs={12} md={6} sx={{ p:2.5, borderRight:{md:`1px solid rgba(23,70,139,0.08)`} }}>
                <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:1 }}>Today</Typography>
                {el.current.length===0 ? <Typography variant="body2" sx={{color:"#9bb0cc"}}>—</Typography> : el.current.map((a,i)=>(
                    <Box key={i} sx={{ py:0.7, borderBottom:i<el.current.length-1?`1px dashed rgba(23,70,139,0.1)`:"none", display:"flex", gap:1 }}><Chip label={`G${a.g}`} size="small" sx={{background:BAT_LIGHT,color:BAT_NAVY,fontWeight:700,height:18,fontSize:"0.65rem",flexShrink:0}}/><Typography variant="body2" sx={{color:"#2a3a50",lineHeight:1.5}}>{a.t}</Typography></Box>
                  ))}
              </Grid>
              <Grid item xs={12} md={6} sx={{ p:2.5 }}>
                <Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", mb:1 }}>2027</Typography>
                {el.future.length===0 ? <Typography variant="body2" sx={{color:"#9bb0cc"}}>—</Typography> : el.future.map((a,i)=>(
                    <Box key={i} sx={{ py:0.7, borderBottom:i<el.future.length-1?`1px dashed rgba(250,180,30,0.2)`:"none", display:"flex", gap:1 }}><Chip label={`G${a.g}`} size="small" sx={{background:"#fffbea",color:"#8a5900",fontWeight:700,height:18,fontSize:"0.65rem",flexShrink:0}}/><Typography variant="body2" sx={{color:"#2a3a50",lineHeight:1.5}}>{a.t}</Typography></Box>
                  ))}
              </Grid>
            </Grid>
          </Paper>
        ))}
      </Stack>
    </Box>
  );
}

// ─── GROUP ────────────────────────────────────────────────────────────────────
function Group({ onExit }) {
  const [joinedSession, setJoinedSession] = useState(localStorage.getItem("grp_session")||null);
  const [joinedGroup,   setJoinedGroup]   = useState(parseInt(localStorage.getItem("grp_num"))||null);
  const [session, setSession] = useState(null);
  const [gd, setGd]           = useState(null);
  const [busy, setBusy]       = useState(false);

  const refresh = useCallback(async()=>{
    if (!joinedSession||!joinedGroup) return;
    const s=await loadSession(joinedSession);
    if (!s) { localStorage.removeItem("grp_session"); localStorage.removeItem("grp_num"); setJoinedSession(null); return; }
    setSession(s); const g=await loadGroup(joinedSession,joinedGroup); if (g) setGd(g);
  },[joinedSession,joinedGroup]);

  useEffect(()=>{ if (!joinedSession) return; refresh(); const t=setInterval(refresh,8000); return()=>clearInterval(t); },[joinedSession,refresh]);

  const onJoin=async(code,n)=>{
    setBusy(true); const s=await loadSession(code);
    if (!s){alert("Session code not found.");setBusy(false);return;}
    if (n<1||n>s.numGroups){alert(`Group number must be between 1 and ${s.numGroups}.`);setBusy(false);return;}
    await saveGroup(code,n,{joined:true});
    localStorage.setItem("grp_session",code); localStorage.setItem("grp_num",n);
    setJoinedSession(code); setJoinedGroup(n); setBusy(false); refresh();
  };
  const handleExit=()=>{ localStorage.removeItem("grp_session"); localStorage.removeItem("grp_num"); setJoinedSession(null); onExit(); };

  if (!joinedSession) return <GJoin onJoin={onJoin} onExit={handleExit} busy={busy}/>;
  if (!session||!gd) return <Box sx={{display:"flex",justifyContent:"center",pt:10}}><CircularProgress sx={{color:BAT_NAVY}}/></Box>;
  return <GWork session={session} gd={gd} n={joinedGroup} refresh={refresh}/>;
}

function GJoin({ onJoin, onExit, busy }) {
  const [code, setCode] = useState(""); const [num,  setNum]  = useState("");
  return (
    <Box>
      <Header />
      <Box sx={{ maxWidth:460, mx:"auto", px:2, pt:5 }}>
        <Button startIcon={<ArrowBackIcon/>} onClick={onExit} sx={{ mb:3, color:BAT_NAVY }}>Back</Button>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Group Check-in</Typography>
        <Typography variant="h4" sx={{ color:BAT_DARK, fontWeight:800, mb:3, mt:0.5 }}>Join the session</Typography>
        <Paper elevation={2} sx={{ p:3.5 }}>
          <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5 }}>Session Code</Typography>
          <TextField fullWidth value={code} onChange={e=>setCode(e.target.value.toUpperCase().slice(0,4))} placeholder="ABCD" variant="outlined" sx={{ mb:2.5, "& input":{ textAlign:"center", letterSpacing:"0.4em", fontWeight:900, fontSize:"1.8rem", py:1.5 } }}/>
          <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5 }}>Group Number</Typography>
          <TextField fullWidth type="number" inputProps={{min:1,max:9}} value={num} onChange={e=>setNum(e.target.value)} placeholder="1, 2, 3…" variant="outlined" sx={{ mb:3, "& input":{ textAlign:"center", fontWeight:900, fontSize:"1.8rem", py:1.5 } }}/>
          <Button fullWidth variant="contained" color="primary" size="large" disabled={!code||code.length<4||!num||busy} onClick={()=>onJoin(code,parseInt(num,10))} endIcon={busy?<CircularProgress size={16} sx={{color:"#fff"}}/>:<NavigateNextIcon/>} sx={{ py:1.5 }}>Join as Group {num||"…"}</Button>
        </Paper>
      </Box>
    </Box>
  );
}

function GWork({ session, gd, n, refresh }) {
  const p = session.currentPhase;
  if (p==="setup")    return <GWait session={session} n={n} msg="Session hasn't started yet. Your facilitator will begin shortly."/>;
  if (p==="complete") return <GWait session={session} n={n} msg="Session wrap-up in progress. Your facilitator is walking through the synthesis."/>;
  if (p==="phase1")   return <GP1 session={session} gd={gd} n={n} refresh={refresh}/>;
  if (p==="phase2")   return <GP2 session={session} gd={gd} n={n} refresh={refresh}/>;
  if (p==="phase3")   return <GP3 session={session} gd={gd} n={n} refresh={refresh}/>;
  if (p==="phase4")   return <GP4 session={session} gd={gd} n={n} refresh={refresh}/>;
  return null;
}

function GWait({ session, n, msg }) {
  return (
    <Box>
      <Header code={session.code} group={n} phase={session.currentPhase}/>
      <Box sx={{ maxWidth:420, mx:"auto", textAlign:"center", pt:8, px:2 }}>
        <Box sx={{ width:64, height:64, borderRadius:"50%", background:BAT_LIGHT, mx:"auto", mb:3, display:"flex", alignItems:"center", justifyContent:"center" }}><CircularProgress size={28} sx={{color:BAT_NAVY}}/></Box>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Group {n}</Typography>
        <Typography variant="h5" sx={{ color:BAT_DARK, fontWeight:800, mt:1, mb:1.5 }}>Hold tight</Typography>
        <Typography sx={{ color:"#4A6080", lineHeight:1.7 }}>{msg}</Typography>
      </Box>
    </Box>
  );
}

function useAutosave(data, save) {
  const ref = useRef(data);
  useEffect(()=>{ ref.current=data; const t=setTimeout(()=>save(ref.current),2000); return()=>clearTimeout(t); },[data,save]);
}

function AutosaveChip({ saving }) { return <Box sx={{ display:"flex", justifyContent:"flex-end", mt:1 }}><Typography variant="caption" sx={{ color:"#9bb0cc" }}>{saving ? "Saving…" : "✓ Autosaved"}</Typography></Box>; }

function OPMCard({ el, value, onChange }) {
  return (
    <Paper elevation={1} sx={{ mb:2, overflow:"hidden" }}>
      <Box sx={{ px:2.5, py:1.5, background:`linear-gradient(90deg,${BAT_LIGHT},#f8fafd)`, borderBottom:`1px solid rgba(23,70,139,0.08)` }}>
        <Typography variant="h6" sx={{ fontWeight:800, color:BAT_DARK, mb:0.2 }}>{el.label}</Typography>
        <Typography variant="caption" sx={{ color:"#4A6080", lineHeight:1.4 }}>{el.def}</Typography>
      </Box>
      <Box sx={{ p:2.5 }}><SmartField label={el.prompt} ctx={`Current state of ${el.label}: ${el.prompt}`} value={value||""} onChange={onChange} rows={2}/></Box>
    </Paper>
  );
}

function GP1({ session, gd, n, refresh }) {
  const [a,setA] = useState(gd.phase1||{}); const [saving,setSaving] = useState(false);
  const save = useCallback(async(d)=>{ setSaving(true); await saveGroup(session.code,n,{phase1:d}); refresh(); setSaving(false); },[session.code,n,refresh]);
  useAutosave(a,save);
  return (
    <Box>
      <Header code={session.code} group={n} phase="phase1"/>
      <Box sx={{ maxWidth:800, mx:"auto", px:2, pt:3, pb:8 }}>
        <PhaseStepper current="phase1"/>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase I</Typography>
        <Typography variant="h4" sx={{ color:BAT_DARK, fontWeight:800, mt:0.5, mb:3 }}>The Smell of the Place</Typography>
        {OPM.map(el=><OPMCard key={el.key} el={el} value={a[el.key]} onChange={v=>setA({...a,[el.key]:v})}/>)}
        <AutosaveChip saving={saving}/>
      </Box>
    </Box>
  );
}

function GP2({ session, gd, n, refresh }) {
  const blank = { headline:"",action1:"",action2:"",action3:"",frontlineQuote:"" };
  const [np,setNp] = useState(gd.phase2?.headline!==undefined?gd.phase2:blank); const [saving,setSaving] = useState(false);
  const save = useCallback(async(d)=>{ setSaving(true); await saveGroup(session.code,n,{phase2:d}); refresh(); setSaving(false); },[session.code,n,refresh]);
  useAutosave(np,save);
  return (
    <Box>
      <Header code={session.code} group={n} phase="phase2"/>
      <Box sx={{ maxWidth:800, mx:"auto", px:2, pt:3, pb:8 }}>
        <PhaseStepper current="phase2"/>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase II</Typography>
        <Typography variant="h4" sx={{ color:BAT_DARK, fontWeight:800, mt:0.5, mb:3 }}>The Newspaper from 2027</Typography>
        <Paper elevation={2} sx={{ overflow:"hidden", border:`2px solid ${BAT_NAVY}22` }}>
          <Box sx={{ background:`linear-gradient(135deg,${BAT_DARK},${BAT_NAVY})`, px:3, py:2, borderBottom:`3px solid ${BAT_GOLD}` }}>
            <Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", mb:0.5 }}>Headline · 2027</Typography>
            <TextField fullWidth value={np.headline} onChange={e=>setNp({...np,headline:e.target.value})} variant="outlined" placeholder="Write the headline BAT has earned…" sx={{ "& .MuiOutlinedInput-root":{ background:"rgba(255,255,255,0.1)", borderRadius:1 }, "& input":{ color:"#fff", fontWeight:800, fontSize:"1.15rem", "::placeholder":{color:"rgba(255,255,255,0.4)"} }, "& fieldset":{ borderColor:"rgba(255,255,255,0.2)" } }}/>
          </Box>
          <Box sx={{ p:3 }}>
            <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:1.5 }}>Three things leadership did differently</Typography>
            {["action1","action2","action3"].map((k,i)=><SmartField key={k} label={`Action ${i+1}`} value={np[k]} onChange={v=>setNp({...np,[k]:v})} placeholder="A specific decision or change…" rows={2}/>)}
            <Divider sx={{ my:2.5 }}/><Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:1.5 }}>Frontline Voice</Typography>
            <SmartField label="Quote from a frontline employee, 2027" value={np.frontlineQuote} onChange={v=>setNp({...np,frontlineQuote:v})} rows={2}/>
          </Box>
        </Paper>
        <AutosaveChip saving={saving}/>
      </Box>
    </Box>
  );
}

function GP3({ session, gd, n, refresh }) {
  const [m,setM] = useState(gd.phase3||{}); const [auto,setAuto] = useState(gd.phase3AutoMapped||null);
  const [gen,setGen] = useState(false); const [reviewed,setReviewed] = useState(!!gd.phase3Reviewed); const [saving,setSaving] = useState(false);

  const generate = async()=>{
    setGen(true); const r=await autoMap(gd.phase2||{}); setAuto(r);
    const seed={...r}; for (const e of OPM) if (m[e.key]) seed[e.key]=m[e.key];
    setM(seed); await saveGroup(session.code,n,{phase3:seed,phase3AutoMapped:r}); refresh(); setGen(false);
  };
  
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(()=>{ if (!auto&&!gd.phase3AutoMapped) generate(); },[]);

  const save=useCallback(async(d)=>{ setSaving(true); await saveGroup(session.code,n,{phase3:d}); refresh(); setSaving(false); },[session.code,n,refresh]);
  useAutosave(m,save);

  const markReviewed=async()=>{ await saveGroup(session.code,n,{phase3Reviewed:true}); setReviewed(true); refresh(); };

  return (
    <Box>
      <Header code={session.code} group={n} phase="phase3"/>
      <Box sx={{ maxWidth:800, mx:"auto", px:2, pt:3, pb:8 }}>
        <PhaseStepper current="phase3"/>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase III</Typography>
        <Typography variant="h4" sx={{ color:BAT_DARK, fontWeight:800, mt:0.5, mb:3 }}>Reverse-engineer the future</Typography>
        {gen ? <Box sx={{ textAlign:"center", py:8 }}><CircularProgress sx={{color:BAT_NAVY}}/><Typography sx={{color:"#4A6080",mt:2}}>AI is mapping your newspaper…</Typography></Box> : (
            <>
              {OPM.map(el=>{
                const aiText = auto?.[el.key]||""; const na = aiText.toLowerCase().includes("not addressed");
                return (
                  <Paper key={el.key} elevation={1} sx={{ mb:2, overflow:"hidden" }}>
                    <Box sx={{ px:2.5, py:1.5, background:`linear-gradient(90deg,${BAT_LIGHT},#f8fafd)`, borderBottom:`1px solid rgba(23,70,139,0.08)`, display:"flex", alignItems:"center", gap:1 }}>
                      <Typography variant="h6" sx={{ fontWeight:800, color:BAT_DARK, flex:1 }}>{el.label}</Typography>
                      {na && <Chip label="Not Addressed" size="small" icon={<WarningAmberIcon sx={{fontSize:"14px !important"}}/>} sx={{ background:"#fff3cd", color:"#856404", borderColor:"#ffc107", border:"1px solid" }}/>}
                    </Box>
                    <Box sx={{ p:2.5 }}>
                      <Typography variant="caption" sx={{ color:"#4A6080", display:"block", mb:1.5, fontStyle:"italic" }}>{el.futurePrompt}</Typography>
                      {aiText && <Alert severity="info" icon={<AutoFixHighIcon fontSize="small"/>} sx={{ mb:1.5, background:"#EEF3FB", border:`1px solid ${BAT_NAVY}22`, "& .MuiAlert-message":{ fontSize:"0.82rem", color:"#1a3a6b", fontStyle:"italic" } }}><strong>AI Suggested:</strong> {aiText}</Alert>}
                      <TextField fullWidth multiline rows={2} value={m[el.key]||""} onChange={e=>setM({...m,[el.key]:e.target.value})} variant="outlined" size="small"/>
                    </Box>
                  </Paper>
                );
              })}
              <Box sx={{ display:"flex", justifyContent:"space-between", alignItems:"center", mt:2 }}>
                <Button variant="outlined" startIcon={<RefreshIcon/>} onClick={generate} disabled={gen} sx={{ color:BAT_NAVY, borderColor:BAT_NAVY }}>Regenerate</Button>
                <Button variant="contained" color={reviewed?"success":"primary"} onClick={markReviewed} disabled={reviewed} startIcon={reviewed?<CheckCircleOutlinedIcon/>:<MarkEmailReadIcon/>}>{reviewed?"Reviewed ✓":"Mark as Reviewed"}</Button>
              </Box>
              <AutosaveChip saving={saving}/>
            </>
          )}
      </Box>
    </Box>
  );
}

function GP4({ session, gd, n, refresh }) {
  const [r,setR] = useState(gd.phase4?.r1!==undefined?gd.phase4:{r1:"",r2:"",r3:""}); const [saving,setSaving] = useState(false);
  const save=useCallback(async(d)=>{ setSaving(true); await saveGroup(session.code,n,{phase4:d}); refresh(); setSaving(false); },[session.code,n,refresh]);
  useAutosave(r,save);
  return (
    <Box>
      <Header code={session.code} group={n} phase="phase4"/>
      <Box sx={{ maxWidth:800, mx:"auto", px:2, pt:3, pb:8 }}>
        <PhaseStepper current="phase4"/>
        <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase IV</Typography>
        <Typography variant="h4" sx={{ color:BAT_DARK, fontWeight:800, mt:0.5, mb:3 }}>Name the Roadblocks</Typography>
        <Paper elevation={1} sx={{ p:3 }}>
          {["r1","r2","r3"].map((k,i)=><SmartField key={k} label={`Roadblock #${i+1}`} value={r[k]} onChange={v=>setR({...r,[k]:v})} rows={2}/>)}
          <AutosaveChip saving={saving}/>
        </Paper>
      </Box>
    </Box>
  );
}

// ─── REPORT ───────────────────────────────────────────────────────────────────
function Report({ session, groups, onClose }) {
  const [generatingPPT, setGeneratingPPT] = useState(false);
  const today = new Date().toLocaleDateString("en-GB",{weekday:"long",month:"long",day:"numeric",year:"numeric"});
  
  const elements = OPM.map(el=>{
    const current=[],future=[];
    for(let i=1;i<=session.numGroups;i++){
      const g=groups[i];
      if(g?.phase1?.[el.key]) current.push({g:i,t:g.phase1[el.key]});
      if(g?.phase3?.[el.key]) future.push({g:i,t:g.phase3[el.key]});
    }
    const addressing=future.filter(f=>!f.t.toLowerCase().includes("not addressed")).length;
    return {...el,current,future,addressing};
  });

  const handlePPTX = async () => {
    setGeneratingPPT(true);
    try {
      // 1. Get AI Themes
      const aiData = await generateThemes(groups);
      
      // 2. Build PPTX
      let pres = new pptxgen();
      pres.layout = "LAYOUT_16x9";
      
      // Slide 1: Title
      let slide1 = pres.addSlide();
      slide1.background = { color: "17468B" };
      slide1.addText("Culture Gap Assessment", { x: 1, y: 2, w: 8, h: 1, fontSize: 44, color: "FFFFFF", bold: true, align: "center" });
      slide1.addText(`${session.sessionName} | ${today}`, { x: 1, y: 3.5, w: 8, h: 1, fontSize: 20, color: "FAB41E", align: "center" });

      // Slide 2: Executive Summary
      let slide2 = pres.addSlide();
      slide2.addText("Executive Summary", { x: 0.5, y: 0.5, w: 8, h: 0.8, fontSize: 28, color: "17468B", bold: true });
      slide2.addText(aiData.executiveSummary, { x: 0.5, y: 1.5, w: 9, h: 1.5, fontSize: 16, color: "333333" });
      slide2.addText("Key Themes & Objectives:", { x: 0.5, y: 3.5, w: 8, h: 0.5, fontSize: 18, color: "17468B", bold: true });
      aiData.themes.forEach((t, i) => {
        slide2.addText(`• ${t}`, { x: 0.8, y: 4.2 + (i * 0.5), w: 8.5, h: 0.5, fontSize: 14, color: "333333" });
      });

      // Slide 3: Group Summaries
      let slide3 = pres.addSlide();
      slide3.addText("Group Summaries (At a Glance)", { x: 0.5, y: 0.5, w: 9, h: 0.8, fontSize: 28, color: "17468B", bold: true });
      let yPos = 1.5;
      for(let i=1; i<=session.numGroups; i++) {
        const g = groups[i];
        if(g && g.summaries && g.summaries.phase3) {
          slide3.addText(`Group ${i}:`, { x: 0.5, y: yPos, w: 1.5, h: 0.5, fontSize: 14, color: "FAB41E", bold: true });
          slide3.addText(g.summaries.phase3, { x: 2.0, y: yPos, w: 7.5, h: 0.5, fontSize: 12, color: "333333" });
          yPos += 0.8;
        }
      }

      await pres.writeFile({ fileName: `Culture_Gap_Report_${session.code}.pptx` });
    } catch (e) { alert("Failed to generate PPTX. Please try again."); }
    setGeneratingPPT(false);
  };

  return (
    <Box>
      <Box className="no-print" sx={{ background:`linear-gradient(135deg,${BAT_DARK},${BAT_NAVY})`, px:{xs:2,md:4}, py:1.5, display:"flex", alignItems:"center", gap:2, borderBottom:`3px solid ${BAT_GOLD}` }}>
        <Button startIcon={<ArrowBackIcon/>} onClick={onClose} sx={{ color:"#fff", "&:hover":{background:"rgba(255,255,255,0.1)"} }}>Dashboard</Button>
        <Box sx={{ flex:1 }}/>
        <Button variant="outlined" sx={{ color: BAT_GOLD, borderColor: BAT_GOLD }} startIcon={generatingPPT ? <CircularProgress size={16} sx={{color:BAT_GOLD}}/> : <DownloadIcon/>} onClick={handlePPTX} disabled={generatingPPT}>
          {generatingPPT ? "Generating PPTX..." : "Generate Executive PPTX"}
        </Button>
        <Button variant="contained" color="secondary" startIcon={<PrintIcon/>} onClick={()=>window.print()}>Print / Save Detailed PDF</Button>
      </Box>

      <Box sx={{ maxWidth:1100, mx:"auto", px:{xs:2,md:4}, pt:4, pb:10 }}>
        <Box sx={{ textAlign:"center", mb:5 }}>
          <Box sx={{ display:"flex", justifyContent:"center", alignItems:"center", gap:3, mb:3 }}>
            <Box component="img" src="/BAT.png" alt="BAT" sx={{ height:40 }} onError={e=>{e.target.style.display="none";}}/>
            <Divider orientation="vertical" flexItem sx={{ borderColor:BAT_GOLD }}/>
            <Box component="img" src="/logo.png" alt="Carnelian Co" sx={{ height:28, opacity:0.7, filter: "brightness(0)" }} onError={e=>{e.target.style.display="none";}}/>
          </Box>
          <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Culture Gap Assessment — Detailed Output</Typography>
          <Typography variant="h3" sx={{ color:BAT_DARK, fontWeight:800, mt:1, mb:1 }}>{session.sessionName}</Typography>
          <Typography sx={{ color:"#4A6080" }}>Session {session.code} · {today} · {session.numGroups} groups</Typography>
        </Box>

        {elements.map(el=>(
          <Box key={el.key} sx={{ mb:4, pb:4, borderBottom:`1px solid rgba(23,70,139,0.12)` }}>
            <Box sx={{ display:"flex", alignItems:"center", mb:2.5, gap:2 }}>
              <Box sx={{ width:6, height:36, background:`linear-gradient(${BAT_NAVY},${BAT_GOLD})`, borderRadius:3, flexShrink:0 }}/>
              <Typography variant="h5" sx={{ fontWeight:800, color:BAT_DARK, flex:1 }}>{el.label}</Typography>
              <Chip label={`${el.addressing}/${session.numGroups} addressed`} size="small" sx={{ background:el.addressing>0?`${BAT_GOLD}22`:BAT_LIGHT, color:BAT_DARK, fontWeight:700 }}/>
            </Box>
            <Grid container spacing={3}>
              <Grid item xs={12} md={6}>
                <Box sx={{ px:2, py:1.5, background:BAT_LIGHT, borderRadius:1, mb:1 }}><Typography variant="overline" sx={{ color:BAT_NAVY, fontWeight:700 }}>Today</Typography></Box>
                {el.current.map((a,i)=>(<Box key={i} sx={{ display:"flex", gap:1.5, py:1, borderBottom:`1px dashed rgba(23,70,139,0.1)` }}><Chip label={`G${a.g}`} size="small" sx={{ background:BAT_LIGHT, color:BAT_NAVY, fontWeight:700, height:20, fontSize:"0.65rem", flexShrink:0 }}/><Typography variant="body2" sx={{ color:"#2a3a50", lineHeight:1.6 }}>{a.t}</Typography></Box>))}
              </Grid>
              <Grid item xs={12} md={6}>
                <Box sx={{ px:2, py:1.5, background:"#fffbea", borderRadius:1, mb:1 }}><Typography variant="overline" sx={{ color:"#8a5900", fontWeight:700 }}>2027</Typography></Box>
                {el.future.map((a,i)=>(<Box key={i} sx={{ display:"flex", gap:1.5, py:1, borderBottom:`1px dashed rgba(250,180,30,0.2)` }}><Chip label={`G${a.g}`} size="small" sx={{ background:"#fffbea", color:"#8a5900", fontWeight:700, height:20, fontSize:"0.65rem", flexShrink:0 }}/><Typography variant="body2" sx={{ color:"#2a3a50", lineHeight:1.6 }}>{a.t}</Typography></Box>))}
              </Grid>
            </Grid>
          </Box>
        ))}
        <Box sx={{ textAlign:"center", pt:3, borderTop:`1px solid rgba(23,70,139,0.12)` }}><Typography variant="caption" sx={{ color:"#9bb0cc" }}>Generated by Carnelian Co · carnelianco.com · Confidential</Typography></Box>
      </Box>
      <style>{`@media print { .no-print { display:none !important; } body { background:#fff; } @page { margin:0.6in; } }`}</style>
    </Box>
  );
}

// ─── ROOT ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [role, setRole] = useState(null);
  return (
    <ThemeProvider theme={theme}>
      <CssBaseline/>
      <style>{`body { background:#F4F6FB; }`}</style>
      {!role      && <Landing go={setRole}/>}
      {role==="fac" && <Facilitator onExit={()=>setRole(null)}/>}
      {role==="grp" && <Group onExit={()=>setRole(null)}/>}
    </ThemeProvider>
  );
}