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

async function generateSummary(data, phase) {
  try {
    const r = await fetch(`${API_URL}/ai/summarize`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({data: JSON.stringify(data), phase}) });
    const d = await r.json(); return d.text;
  } catch { return "Summary unavailable."; }
}
async function generateThemes(data) {
  try {
    const r = await fetch(`${API_URL}/ai/themes`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({data: JSON.stringify(data)}) });
    const d = await r.json(); 
    
    // Safely extract JSON even if Claude adds conversational text before/after it
    const text = d.text;
    const start = text.indexOf('{');
    const end = text.lastIndexOf('}');
    if (start !== -1 && end !== -1) {
      return JSON.parse(text.substring(start, end + 1));
    }
    return JSON.parse(text.replace(/```json|```/g,"").trim());
  } catch (e) { 
    console.error("JSON Parse Error:", e);
    return { executiveSummary: "Summary unavailable.", themes: ["Theme 1", "Theme 2", "Theme 3"] }; 
  }
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
    const text = await generateSummary(group[phaseKey], phaseKey);
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
                {summary ? (
  <Box sx={{ color: "#4A6080", fontSize: "0.75rem", lineHeight: 1.6, "& strong": { color: BAT_DARK, fontWeight: 700 }, "& ul": { pl: 2, mt: 0.5 }, "& li": { mb: 0.3 } }}
    dangerouslySetInnerHTML={{ __html: (() => {
      let html = summary;
      html = html.replace(/(\|.+\|\n?)+/g, "");
      html = html.replace(/^---+$/gm, "<hr/>");
      html = html.replace(/^#{1,2} (.+)$/gm, "<strong style='display:block;font-size:0.8rem;color:#17468B;margin-top:8px'>$1</strong>");
      html = html.replace(/^#{3,4} (.+)$/gm, "<strong style='display:block;font-size:0.78rem;color:#0d2a5a;margin-top:6px'>$1</strong>");
      html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
      html = html.replace(/^> (.+)$/gm, "<blockquote style='border-left:3px solid #FAB41E;padding-left:10px;margin:6px 0;font-style:italic;color:#0d2a5a'>$1</blockquote>");
      html = html.replace(/(^[•\-\*] .+$(\n[•\-\*] .+$)*)/gm, (match) => {
        const items = match.split("\n").filter(l => l.trim()).map(l => `<li>${l.replace(/^[•\-\*] /, "")}</li>`).join("");
        return `<ul style='padding-left:16px;margin:4px 0'>${items}</ul>`;
      });
      html = html.replace(/\n{2,}/g, "<br/>");
      html = html.replace(/\n/g, " ");
      return html;
    })() }}
  />
) : <Typography variant="body2" sx={{ color: "#9bb0cc", fontSize: "0.7rem", fontStyle: "italic" }}>No summary yet.</Typography>}</Box>
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
    const aiData = await generateThemes(groups);
    let pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";

    const NAVY  = "17468B";
    const DARK  = "0d2a5a";
    const GOLD  = "FAB41E";
    const WHITE = "FFFFFF";
    const LIGHT = "EEF3FB";
    const GREY  = "4A6080";
    const mkShadow = () => ({ type:"outer", blur:6, offset:2, angle:135, color:"000000", opacity:0.10 });

    // helper — strip markdown to plain text
    const plain = (t="") => t
      .replace(/\*\*(.+?)\*\*/g, "$1")
      .replace(/^#{1,4} /gm, "")
      .replace(/^[>•\-\*] /gm, "")
      .replace(/\|.+\|/g, "")
      .replace(/---+/g, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim();

    // helper — extract bullet lines from AI text
    const HEADER_PATTERN = /^(Current State Summary|Common Themes|Problematic Behaviours|Key Cultural Findings|Transformation Priorities|Bottom Line|Headline Insight|What This Group|Implied Actionable|Cultural & Behaviour|Facilitator Note|Phase \d|Roadblock Summary|Top Roadblocks|Commitments Made|Full Arc|Gap Summary|Key Gaps|Behavioural Shifts|Connection to|Biggest Risk|Recommended Priority|Summary|Key Insights|Actionable Steps|Critical Issues)/i;

    const bullets = (t="", max=6) => {
      return t.split("\n")
        .map(l => l
          .replace(/^#{1,6}\s*/,"")
          .replace(/^[•\-\*\d\.]+\s*/,"")
          .replace(/\*\*(.+?)\*\*/g,"$1")
          .replace(/\*(.+?)\*/g,"$1")
          .replace(/^>\s*/,"")
          .replace(/\s*\*+\s*$/,"")
          .trim()
        )
        .filter(l =>
          l.length > 20 &&
          l.length < 220 &&
          !l.match(/^#+/) &&
          !l.includes("---") &&
          !l.includes("|") &&
          !HEADER_PATTERN.test(l) &&
          !l.endsWith("**") &&
          !l.match(/^[A-Z][a-z]+ \d+ (Analysis|Summary|Exercise)/)
        )
        .slice(0, max);
    };
/////////
    // ── SLIDE 1: Title ──────────────────────────────────────────────────────
    let s1 = pres.addSlide();
    s1.background = { color: DARK };
    s1.addShape(pres.shapes.RECTANGLE, { x:0, y:4.72, w:10, h:0.08, fill:{ color:GOLD }, line:{ color:GOLD } });
    s1.addShape(pres.shapes.RECTANGLE, { x:0, y:4.8,  w:10, h:0.83, fill:{ color:NAVY }, line:{ color:NAVY } });
    s1.addText("CULTURE GAP ASSESSMENT", { x:0.6, y:1.0, w:8.8, h:0.5, fontSize:12, bold:true, color:GOLD, charSpacing:5 });
    s1.addText(session.sessionName, { x:0.6, y:1.55, w:8.8, h:1.3, fontSize:40, bold:true, color:WHITE, fontFace:"Calibri" });
    s1.addText(`${today}  ·  Session ${session.code}  ·  ${session.numGroups} Groups`, { x:0.6, y:3.0, w:8.8, h:0.45, fontSize:13, color:"8aaad4" });
    s1.addText("Powered by Carnelian Co  ·  Confidential", { x:0.6, y:4.88, w:8.8, h:0.4, fontSize:10, color:"aabbdd", italic:true });

    // ── SLIDE 2: Executive Summary (concise) ────────────────────────────────
    let s2 = pres.addSlide();
    s2.background = { color:"F4F6FB" };
    s2.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.9, fill:{ color:NAVY }, line:{ color:NAVY } });
    s2.addShape(pres.shapes.RECTANGLE, { x:0, y:0.9, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
    s2.addText("EXECUTIVE SUMMARY", { x:0.5, y:0.18, w:9, h:0.55, fontSize:20, bold:true, color:WHITE });

    // Trim executive summary to max 3 sentences
    const execRaw = plain(aiData.executiveSummary || "");
    const execSentences = execRaw.match(/[^.!?]+[.!?]+/g) || [execRaw];
    const execShort = execSentences.slice(0,3).join(" ").trim();

    s2.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.1, w:5.8, h:4.1, fill:{ color:WHITE }, line:{ color:"dde5f0" }, shadow:mkShadow() });
    s2.addText(execShort, { x:0.6, y:1.25, w:5.4, h:1.4, fontSize:12, color:GREY, wrap:true, valign:"top" });

    // Key issues bullets from summary
    const issueLines = bullets(aiData.executiveSummary || "", 4);
    if (issueLines.length) {
      s2.addText("KEY FINDINGS", { x:0.6, y:2.75, w:5.4, h:0.3, fontSize:9, bold:true, color:NAVY, charSpacing:2 });
      const issueItems = issueLines.map((l,idx) => ({
        text: l, options:{ bullet:true, breakLine: idx < issueLines.length-1, fontSize:11, color:GREY, paraSpaceAfter:4 }
      }));
      s2.addText(issueItems, { x:0.6, y:3.1, w:5.4, h:1.8 });
    }

    // Themes column
    const themes = (aiData.themes || []).slice(0,3);
    s2.addShape(pres.shapes.RECTANGLE, { x:6.4, y:1.1, w:3.2, h:4.1, fill:{ color:NAVY }, line:{ color:NAVY }, shadow:mkShadow() });
    s2.addText("KEY THEMES", { x:6.5, y:1.22, w:3.0, h:0.3, fontSize:9, bold:true, color:GOLD, charSpacing:3 });
    themes.forEach((t, i) => {
      const tPlain = plain(t);
      s2.addShape(pres.shapes.RECTANGLE, { x:6.5, y:1.62+i*1.18, w:3.0, h:1.08, fill:{ color:"1e56a8" }, line:{ color:"1e56a8" } });
      s2.addText(`${i+1}`, { x:6.55, y:1.65+i*1.18, w:0.3, h:0.28, fontSize:9, bold:true, color:GOLD });
      s2.addText(tPlain, { x:6.85, y:1.63+i*1.18, w:2.55, h:1.04, fontSize:9, color:WHITE, wrap:true, valign:"middle", shrinkText:true });
    });

    // ── SLIDE 3: Behavioural Shifts ─────────────────────────────────────────
    const bShifts = aiData.behaviouralShifts || [];
    if (bShifts.length) {
      let s3 = pres.addSlide();
      s3.background = { color:"F4F6FB" };
      s3.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.9, fill:{ color:DARK }, line:{ color:DARK } });
      s3.addShape(pres.shapes.RECTANGLE, { x:0, y:0.9, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
      s3.addText("BEHAVIOURAL SHIFTS REQUIRED", { x:0.5, y:0.18, w:9, h:0.55, fontSize:20, bold:true, color:WHITE });
      bShifts.slice(0,4).forEach((shift, i) => {
        const sp = plain(shift);
        // Try → first, then fall back to " to " split
        let fromPart = "", toPart = "";
        if (sp.includes("→")) {
          const parts = sp.split("→");
          fromPart = parts[0].replace(/^From\s*/i,"").trim();
          toPart   = parts[1]?.replace(/^To\s*/i,"").replace(/^:\s*/,"").trim() || "";
        } else {
          // match "from X to Y" pattern case-insensitively
          const m = sp.match(/^From\s+(.+?)\s+to\s+(.+?)(?:\s*[:\-—].*)?$/i);
          if (m) { fromPart = m[1].trim(); toPart = m[2].trim(); }
          else { fromPart = sp; toPart = ""; }
        }
        // strip trailing colon/reason from toPart for display
        toPart = toPart.replace(/\s*[:\-—].*$/,"").trim();
        const yy = 1.05 + i*1.1;
        s3.addShape(pres.shapes.RECTANGLE, { x:0.4, y:yy, w:9.2, h:0.95, fill:{ color:WHITE }, line:{ color:"dde5f0" }, shadow:mkShadow() });
        s3.addShape(pres.shapes.RECTANGLE, { x:0.4, y:yy, w:0.06, h:0.95, fill:{ color:GOLD }, line:{ color:GOLD } });
        s3.addText(`${i+1}`, { x:0.55, y:yy+0.08, w:0.4, h:0.35, fontSize:14, bold:true, color:NAVY });
        s3.addText(`FROM:  ${fromPart}`, { x:1.05, y:yy+0.06, w:8.3, h:0.32, fontSize:11, color:"C0392B", wrap:true, shrinkText:true });
        if (toPart) s3.addText(`TO:  ${toPart}`, { x:1.05, y:yy+0.52, w:8.3, h:0.32, fontSize:11, color:"166534", wrap:true, shrinkText:true });
      });
    }

    // ── SLIDE 4: Priority Actions ───────────────────────────────────────────
    const pActions = aiData.priorityActions || [];
    if (pActions.length) {
      let s4 = pres.addSlide();
      s4.background = { color:"F4F6FB" };
      s4.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.9, fill:{ color:NAVY }, line:{ color:NAVY } });
      s4.addShape(pres.shapes.RECTANGLE, { x:0, y:0.9, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
      s4.addText("PRIORITY ACTIONS", { x:0.5, y:0.18, w:9, h:0.55, fontSize:20, bold:true, color:WHITE });
      const aColors = [GOLD, "27a060", NAVY];
      pActions.slice(0,3).forEach((action, i) => {
        const ap = plain(action);
        // Split label from reason at " — "
        const dashIdx = ap.indexOf(" — ");
        const label  = dashIdx > -1 ? ap.substring(0, dashIdx) : ap;
        const reason = dashIdx > -1 ? ap.substring(dashIdx+3) : "";
        
        const yy = 1.05 + i*1.45;
        s4.addShape(pres.shapes.RECTANGLE, { x:0.4, y:yy, w:1.0, h:1.25, fill:{ color:aColors[i]||NAVY }, line:{ color:aColors[i]||NAVY } });
        s4.addText(`0${i+1}`, { x:0.4, y:yy+0.35, w:1.0, h:0.55, fontSize:22, bold:true, color:WHITE, align:"center" });
        s4.addShape(pres.shapes.RECTANGLE, { x:1.5, y:yy, w:8.1, h:1.25, fill:{ color:WHITE }, line:{ color:"dde5f0" }, shadow:mkShadow() });
        
        // Removed the "Short" variables and slightly increased the 'h' (height) to allow for text wrapping
        s4.addText(label, { x:1.65, y:yy+0.1, w:7.8, h:0.45, fontSize:12, bold:true, color:DARK, wrap:true });
        if (reason) s4.addText(reason, { x:1.65, y:yy+0.55, w:7.8, h:0.65, fontSize:10, color:GREY, wrap:true });
      });
    }

    // ── SLIDE 5: Group Insights ─────────────────────────────────────────────
    if (aiData.groupInsights) {
      let s5 = pres.addSlide();
      s5.background = { color:"F4F6FB" };
      s5.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.9, fill:{ color:DARK }, line:{ color:DARK } });
      s5.addShape(pres.shapes.RECTANGLE, { x:0, y:0.9, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
      s5.addText("GROUP INSIGHTS", { x:0.5, y:0.18, w:9, h:0.55, fontSize:20, bold:true, color:WHITE });

      const commonBullets  = bullets(aiData.groupInsights.common || "", 4);
      const divergeBullets = bullets(aiData.groupInsights.divergence || "", 4);

      const commonText    = plain(aiData.groupInsights?.common || "");
      const divergeText   = plain(aiData.groupInsights?.divergence || "");
      // detect if AI returned same content for both — if so show a note on diverge side
      const isSame = commonText.trim().toLowerCase() === divergeText.trim().toLowerCase();

      [[commonBullets, commonText, "WHAT ALL GROUPS AGREED", 0.4, NAVY],
       [divergeBullets, isSame ? "" : divergeText, "WHERE GROUPS DIVERGED", 5.2, DARK]].forEach(([blist, fallbackText, title, xPos, bgCol], panelIdx) => {
        s5.addShape(pres.shapes.RECTANGLE, { x:xPos, y:1.1, w:4.5, h:4.1, fill:{ color:WHITE }, line:{ color:"dde5f0" }, shadow:mkShadow() });
        s5.addShape(pres.shapes.RECTANGLE, { x:xPos, y:1.1, w:4.5, h:0.45, fill:{ color:bgCol }, line:{ color:bgCol } });
        s5.addText(title, { x:xPos+0.15, y:1.17, w:4.2, h:0.3, fontSize:8, bold:true, color:GOLD, charSpacing:2 });
        if (panelIdx === 1 && isSame) {
          // AI gave same text — write a meaningful note instead
          const noteLines = [];
          for (let n=1; n<=session.numGroups; n++) {
            const g = groups[n];
            if (!g) continue;
            const rb = [g.phase4?.r1, g.phase4?.r2, g.phase4?.r3].filter(Boolean);
            if (rb.length) noteLines.push(`Group ${n} flagged: ${rb[0]}`);
          }
          const noteItems = noteLines.length
            ? noteLines.map((l,idx) => ({ text:l, options:{ bullet:true, breakLine:idx<noteLines.length-1, fontSize:11, color:GREY, paraSpaceAfter:6 } }))
            : [{ text:"Groups showed high alignment — no significant divergence detected in this session.", options:{ fontSize:11, color:GREY } }];
          s5.addText(noteItems, { x:xPos+0.2, y:1.65, w:4.1, h:3.4 });
        } else if (blist.length) {
          const items = blist.map((l,idx) => ({ text:l, options:{ bullet:true, breakLine:idx<blist.length-1, fontSize:11, color:GREY, paraSpaceAfter:6 } }));
          s5.addText(items, { x:xPos+0.2, y:1.65, w:4.1, h:3.4 });
        } else if (fallbackText) {
          s5.addText(fallbackText, { x:xPos+0.2, y:1.65, w:4.1, h:3.4, fontSize:11, color:GREY, wrap:true });
        }
      });
    }

    // ── NEW SLIDES: Management & Contributor Shifts ─────────────────────────
    const levels = [
      { key: "seniorManagement", title: "SENIOR MANAGEMENT COMMITMENTS", bg: DARK },
      { key: "middleManagement", title: "MIDDLE MANAGEMENT SHIFTS", bg: NAVY },
      { key: "individualContributors", title: "INDIVIDUAL CONTRIBUTOR SHIFTS", bg: "1e56a8" }
    ];

    levels.forEach(lvl => {
      const data = aiData[lvl.key];
      if (!data) return;

      let sLvl = pres.addSlide();
      sLvl.background = { color: "F4F6FB" };
      sLvl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: lvl.bg }, line: { color: lvl.bg } });
      sLvl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.9, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });
      sLvl.addText(lvl.title, { x: 0.5, y: 0.18, w: 9, h: 0.55, fontSize: 20, bold: true, color: WHITE });

      // Context Box
      sLvl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 9.2, h: 0.8, fill: { color: WHITE }, line: { color: "dde5f0" }, shadow: mkShadow() });
      sLvl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 0.06, h: 0.8, fill: { color: GOLD }, line: { color: GOLD } });
      sLvl.addText("CONTEXT", { x: 0.55, y: 1.15, w: 2, h: 0.2, fontSize: 8, bold: true, color: NAVY, charSpacing: 2 });
      sLvl.addText(plain(data.context || ""), { x: 0.55, y: 1.35, w: 8.9, h: 0.5, fontSize: 10, color: GREY, wrap: true, italic: true });

      // Themes (Left)
      sLvl.addText("KEY THEMES", { x: 0.4, y: 2.05, w: 4.4, h: 0.25, fontSize: 9, bold: true, color: NAVY, charSpacing: 2 });
      const themeItems = (data.themes || []).map((t, i) => ({ text: plain(t), options: { bullet: true, breakLine: i < data.themes.length - 1, fontSize: 10, color: GREY, paraSpaceAfter: 6 } }));
      sLvl.addText(themeItems, { x: 0.4, y: 2.3, w: 4.4, h: 1.5 });

      // Behaviours (Right)
      sLvl.addText("CRITICAL BEHAVIOURS", { x: 5.2, y: 2.05, w: 4.4, h: 0.25, fontSize: 9, bold: true, color: NAVY, charSpacing: 2 });
      const behItems = (data.behaviours || []).map((b, i) => {
        const bp = plain(b);
        const colonIdx = bp.indexOf(":");
        if (colonIdx > -1) {
          return [
            { text: bp.substring(0, colonIdx + 1) + " ", options: { bold: true, fontSize: 10, color: DARK } },
            { text: bp.substring(colonIdx + 1), options: { breakLine: i < data.behaviours.length - 1, fontSize: 10, color: GREY, paraSpaceAfter: 6 } }
          ];
        }
        return { text: bp, options: { bullet: true, breakLine: i < data.behaviours.length - 1, fontSize: 10, color: GREY, paraSpaceAfter: 6 } };
      }).flat();
      sLvl.addText(behItems, { x: 5.2, y: 2.3, w: 4.4, h: 1.5 });

      // Actions (Bottom)
      sLvl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.9, w: 9.2, h: 1.3, fill: { color: "fffbea" }, line: { color: "f0e0a0" } });
      sLvl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.9, w: 9.2, h: 0.25, fill: { color: GOLD }, line: { color: GOLD } });
      sLvl.addText("REQUIRED ACTIONS", { x: 0.5, y: 3.92, w: 4, h: 0.2, fontSize: 8, bold: true, color: DARK, charSpacing: 2 });
      const actItems = (data.actions || []).map((a, i) => ({ text: plain(a), options: { bullet: true, breakLine: i < data.actions.length - 1, fontSize: 10, color: "5a4000", paraSpaceAfter: 4 } }));
      sLvl.addText(actItems, { x: 0.5, y: 4.2, w: 9.0, h: 0.9 });
    });
    s6.background = { color:"F4F6FB" };
    s6.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.9, fill:{ color:NAVY }, line:{ color:NAVY } });
    s6.addShape(pres.shapes.RECTANGLE, { x:0, y:0.9, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
    s6.addText("THE TRANSFORMATION GAP — 6 DESIGN ELEMENTS", { x:0.5, y:0.18, w:9, h:0.55, fontSize:18, bold:true, color:WHITE });

    // Build a combined today/future string per element for AI to summarise
    const gapSummaries = {};
    for (const el of OPM) {
      const todayLines = [], futureLines = [];
      for (let n=1; n<=session.numGroups; n++) {
        const tv = groups[n]?.phase1?.[el.key]; if (tv) todayLines.push(`G${n}: ${tv}`);
        const fv = groups[n]?.phase3?.[el.key]; if (fv && !fv.toLowerCase().includes("not addressed")) futureLines.push(`G${n}: ${fv}`);
      }
      if (!todayLines.length && !futureLines.length) { gapSummaries[el.key] = { today:"No data", future:"No data", gap:"No data provided." }; continue; }
      try {
        const prompt = `Element: ${el.label}\nTODAY (problems): ${todayLines.join(" | ")}\n2027 VISION: ${futureLines.join(" | ")}\n\nIn exactly 2 short bullet points: what is the core gap between today and 2027 for this element? Each bullet max 15 words. No headers, no preamble, just the 2 bullets starting with •`;
        const r = await fetch(`${API_URL}/ai/nudge`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ system:"You are a concise organisational analyst. Return only bullet points, no other text.", user: prompt }) });
        const d = await r.json();
        gapSummaries[el.key] = d.text || "Gap analysis unavailable.";
      } catch { gapSummaries[el.key] = "Gap analysis unavailable."; }
    }

    OPM.forEach((el, i) => {
      const yy = 1.05 + i * 0.75;
      const gapText = typeof gapSummaries[el.key] === "string"
        ? gapSummaries[el.key].replace(/^[•\-\*] /gm,"").trim()
        : "No data.";
      const gapLines = gapText.split("\n").filter(l=>l.trim()).slice(0,2);

      s6.addShape(pres.shapes.RECTANGLE, { x:0.4, y:yy, w:2.2, h:0.62, fill:{ color:NAVY }, line:{ color:NAVY } });
      s6.addText(el.label, { x:0.5, y:yy+0.18, w:2.0, h:0.28, fontSize:11, bold:true, color:WHITE, align:"center" });

      s6.addShape(pres.shapes.RECTANGLE, { x:2.7, y:yy, w:7.0, h:0.62, fill:{ color:WHITE }, line:{ color:"dde5f0" } });
      s6.addShape(pres.shapes.RECTANGLE, { x:2.7, y:yy, w:0.05, h:0.62, fill:{ color:GOLD }, line:{ color:GOLD } });
      const gapItems = gapLines.map((l,idx)=>({ text:l.replace(/^[•\-\*]\s*/,""), options:{ bullet:true, breakLine:idx<gapLines.length-1, fontSize:10, color:GREY, paraSpaceAfter:2 } }));
      s6.addText(gapItems.length ? gapItems : [{ text:"Gap data not available", options:{ fontSize:10, color:GREY } }], { x:2.85, y:yy+0.06, w:6.7, h:0.52 });
    });

    // ── SLIDES 7+: Per-group AI summary slides ──────────────────────────────
    for (let i = 1; i <= session.numGroups; i++) {
      const g = groups[i];
      if (!g) continue;
      const hasAnyData = (g.phase1 && Object.values(g.phase1).some(v=>v)) ||
                         (g.phase2 && g.phase2.headline) ||
                         (g.phase3 && Object.values(g.phase3).some(v=>v)) ||
                         (g.phase4 && Object.values(g.phase4).some(v=>v));
      if (!hasAnyData) continue;

      // Generate any missing summaries on the fly
      const phasesToGen = ["phase1","phase2","phase3","phase4"];
      for (const pk of phasesToGen) {
        if (!g.summaries?.[pk] && g[pk] && Object.values(g[pk]).some(v=>v)) {
          try {
            const r = await fetch(`${API_URL}/ai/summarize`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ data: JSON.stringify(g[pk]), phase: pk }) });
            const d = await r.json();
            if (!g.summaries) g.summaries = {};
            g.summaries[pk] = d.text || "";
          } catch { g.summaries = g.summaries || {}; g.summaries[pk] = ""; }
        }
      }

      const hasSummaries = g.summaries && Object.values(g.summaries).some(v=>v);
      if (!hasSummaries) continue;

      // Group cover
      let gc = pres.addSlide();
      gc.background = { color:NAVY };
      gc.addShape(pres.shapes.RECTANGLE, { x:0, y:4.72, w:10, h:0.08, fill:{ color:GOLD }, line:{ color:GOLD } });
      gc.addShape(pres.shapes.RECTANGLE, { x:0, y:4.8,  w:10, h:0.83, fill:{ color:DARK }, line:{ color:DARK } });
      gc.addText(`GROUP ${i}`, { x:0.6, y:1.4, w:8.8, h:0.5, fontSize:12, bold:true, color:GOLD, charSpacing:6 });
      gc.addText("AI Analysis & Insights", { x:0.6, y:2.0, w:8.8, h:0.9, fontSize:34, bold:true, color:WHITE });
      gc.addText(session.sessionName, { x:0.6, y:3.1, w:8.8, h:0.4, fontSize:13, color:"8aaad4" });

      // One slide per phase that has a summary
      const phaseMeta = [
        { key:"phase1", label:"Phase 1 — Current State",      bg:NAVY },
        { key:"phase2", label:"Phase 2 — Newspaper 2027",     bg:DARK },
        { key:"phase3", label:"Phase 3 — Vision Mapping",     bg:NAVY },
        { key:"phase4", label:"Phase 4 — Roadblocks",         bg:"8B1A1A" },
      ];

      phaseMeta.forEach(pm => {
        const raw = g.summaries?.[pm.key];
        if (!raw) return;
        const blist = bullets(raw, 7);
        const summary3 = (() => {
          const cleaned = plain(raw)
            .split("\n")
            .filter(l => !HEADER_PATTERN.test(l.trim()) && l.trim().length > 20)
            .join(" ");
          const sentences = cleaned.match(/[^.!?]+[.!?]+/g) || [];
          return sentences.slice(0,2).join(" ").trim();
        })();

        let ps = pres.addSlide();
        ps.background = { color:"F4F6FB" };
        ps.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.9, fill:{ color:pm.bg }, line:{ color:pm.bg } });
        ps.addShape(pres.shapes.RECTANGLE, { x:0, y:0.9, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
        ps.addText(`GROUP ${i}  ·  ${pm.label}`, { x:0.5, y:0.18, w:9, h:0.55, fontSize:18, bold:true, color:WHITE });

        // Summary box
        ps.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.05, w:9.2, h:0.85, fill:{ color:WHITE }, line:{ color:"dde5f0" }, shadow:mkShadow() });
        ps.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.05, w:0.06, h:0.85, fill:{ color:GOLD }, line:{ color:GOLD } });
        ps.addText("SUMMARY", { x:0.58, y:1.08, w:1.5, h:0.22, fontSize:8, bold:true, color:NAVY, charSpacing:2 });
        ps.addText(summary3 || plain(raw).substring(0,180), { x:0.58, y:1.28, w:8.9, h:0.55, fontSize:11, color:GREY, wrap:true, italic:true });

        // Parse sections dynamically based on AI headers
        const sections = [];
        let currentHeader = "KEY INSIGHTS";
        let currentBullets = [];
        
        const lines = plain(raw).split("\n").map(l => l.trim()).filter(l => l);
        let inSummary = false;
        
        lines.forEach(line => {
          if (line === "SUMMARY") { inSummary = true; return; }
          
          // Detect headers (All caps, short, no punctuation at start)
          if (/^[A-Z\s]+$/.test(line) && line.length > 3 && line.length < 35) {
            inSummary = false;
            if (currentBullets.length > 0) {
              sections.push({ header: currentHeader, bullets: currentBullets });
              currentBullets = [];
            }
            currentHeader = line;
          } else if (!inSummary) {
            // Clean up bullet points
            const cleanLine = line.replace(/^[-\*•\d\.]+\s*/, "").trim();
            if (cleanLine.length > 10) currentBullets.push(cleanLine);
          }
        });
        if (currentBullets.length > 0) {
          sections.push({ header: currentHeader, bullets: currentBullets });
        }

        // Render up to 2 columns
        if (sections.length > 0) {
          // Left column
          const leftSec = sections[0];
          ps.addText(leftSec.header, { x:0.5, y:2.05, w:4.5, h:0.28, fontSize:9, bold:true, color:NAVY, charSpacing:2 });
          const leftItems = leftSec.bullets.slice(0,4).map((l,idx)=>({
            text:l, options:{ bullet:true, breakLine:idx<3, fontSize:11, color:GREY, paraSpaceAfter:6 }
          }));
          ps.addText(leftItems, { x:0.5, y:2.38, w:4.5, h:2.8 });

          // Right column
          if (sections.length > 1) {
            const rightSec = sections[1];
            ps.addShape(pres.shapes.RECTANGLE, { x:5.2, y:2.0, w:4.4, h:3.3, fill:{ color:"fffbea" }, line:{ color:"f0e0a0" } });
            ps.addShape(pres.shapes.RECTANGLE, { x:5.2, y:2.0, w:4.4, h:0.35, fill:{ color:GOLD }, line:{ color:GOLD } });
            ps.addText(rightSec.header, { x:5.3, y:2.06, w:4.2, h:0.22, fontSize:8, bold:true, color:DARK, charSpacing:2 });
            const rightItems = rightSec.bullets.slice(0,4).map((l,idx)=>({
              text:l, options:{ bullet:true, breakLine:idx<3, fontSize:11, color:"5a4000", paraSpaceAfter:6 }
            }));
            ps.addText(rightItems, { x:5.3, y:2.42, w:4.2, h:2.7 });
          }
        }
      });
    }

    // ── Final slide ─────────────────────────────────────────────────────────
    let sf = pres.addSlide();
    sf.background = { color:DARK };
    sf.addShape(pres.shapes.RECTANGLE, { x:0, y:2.55, w:10, h:0.06, fill:{ color:GOLD }, line:{ color:GOLD } });
    sf.addText("Thank You", { x:0.6, y:1.1, w:8.8, h:1.1, fontSize:44, bold:true, color:WHITE, align:"center" });
    sf.addText("This document is confidential and prepared exclusively for internal use.", { x:1, y:2.85, w:8, h:0.5, fontSize:11, color:"8aaad4", align:"center", italic:true });
    sf.addText("Carnelian Co  ·  carnelianco.com", { x:1, y:3.55, w:8, h:0.45, fontSize:12, color:GOLD, align:"center" });

    await pres.writeFile({ fileName: `Culture_Gap_Report_${session.code}.pptx` });
  } catch (e) {
    console.error("PPTX Error:", e);
    alert("Failed to generate PPTX. Please try again.");
  }
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

        {/* ── OPM Gap View (existing) ── */}
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

        {/* ── Full Group Responses by Phase ── */}
        <Box sx={{ mt:6, mb:2 }}>
          <Box sx={{ display:"flex", alignItems:"center", gap:2, mb:1 }}>
            <Box sx={{ width:6, height:36, background:`linear-gradient(${BAT_NAVY},${BAT_GOLD})`, borderRadius:3 }}/>
            <Typography variant="h4" sx={{ fontWeight:800, color:BAT_DARK }}>Full Group Responses</Typography>
          </Box>
          <Typography variant="body2" sx={{ color:"#4A6080", mb:3 }}>Every group · every phase · complete input</Typography>
        </Box>

        {Array.from({ length: session.numGroups }, (_, i) => i + 1).map(n => {
          const g = groups[n];
          if (!g) return null;
          return (
            <Box key={n} sx={{ mb:5, pb:5, borderBottom:`2px solid rgba(23,70,139,0.12)` }}>
              {/* Group Header */}
              <Box sx={{ background:`linear-gradient(135deg,${BAT_DARK},${BAT_NAVY})`, borderRadius:1, px:3, py:2, mb:3, display:"flex", alignItems:"center", gap:2 }}>
                <Box sx={{ width:44, height:44, borderRadius:1, background:BAT_GOLD, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                  <Typography sx={{ fontWeight:900, fontSize:"1.2rem", color:BAT_DARK }}>G{n}</Typography>
                </Box>
                <Box>
                  <Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", lineHeight:1 }}>Group {n}</Typography>
                  <Typography variant="h6" sx={{ color:"#fff", fontWeight:800 }}>{session.sessionName}</Typography>
                </Box>
              </Box>

              {/* Phase 1 */}
              {g.phase1 && Object.values(g.phase1).some(v=>v) && (
                <Box sx={{ mb:3 }}>
                  <Box sx={{ px:2, py:1, background:BAT_NAVY, borderRadius:1, mb:1.5, display:"inline-block" }}>
                    <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase 1 — Current State</Typography>
                  </Box>
                  <Grid container spacing={1.5}>
                    {OPM.map(el => g.phase1[el.key] ? (
                      <Grid item xs={12} sm={6} key={el.key}>
                        <Box sx={{ p:2, border:`1px solid rgba(23,70,139,0.12)`, borderRadius:1, height:"100%", background:"#fff" }}>
                          <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5, fontSize:"0.6rem" }}>{el.label}</Typography>
                          <Typography variant="body2" sx={{ color:"#2a3a50", lineHeight:1.6 }}>{g.phase1[el.key]}</Typography>
                        </Box>
                      </Grid>
                    ) : null)}
                  </Grid>
                  {g.summaries?.phase1 && (
                    <Box sx={{ mt:1.5, p:2, background:BAT_LIGHT, borderLeft:`3px solid ${BAT_NAVY}`, borderRadius:1 }}>
                      <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5, fontSize:"0.6rem" }}>AI Summary</Typography>
                      <Box sx={{
  color:"#2a3a50", fontSize:"0.82rem", lineHeight:1.6,
  "& strong":{ color:BAT_DARK, fontWeight:700 },
  "& ul":{ pl:2, mt:0.5, mb:0.5 },
  "& li":{ mb:0.3 },
  "& hr":{ border:"none", borderTop:`1px solid rgba(23,70,139,0.12)`, my:1 },
  "& blockquote":{ borderLeft:`3px solid ${BAT_GOLD}`, pl:1.5, ml:0, fontStyle:"italic", my:1 }
}} dangerouslySetInnerHTML={{ __html: (() => {
  let html = g.summaries.phase1; // change phase key here for each one
  html = html.replace(/(\|.+\|\n?)+/g,"");
  html = html.replace(/^---+$/gm,"<hr/>");
  html = html.replace(/^#{1,2} (.+)$/gm,"<strong style='display:block;font-size:0.8rem;color:#17468B;margin-top:8px'>$1</strong>");
  html = html.replace(/^#{3,4} (.+)$/gm,"<strong style='display:block;font-size:0.78rem;color:#0d2a5a;margin-top:6px'>$1</strong>");
  html = html.replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>");
  html = html.replace(/^> (.+)$/gm,"<blockquote>$1</blockquote>");
  html = html.replace(/(^[•\-\*] .+$(\n[•\-\*] .+$)*)/gm,(match)=>{
    const items = match.split("\n").filter(l=>l.trim()).map(l=>`<li>${l.replace(/^[•\-\*] /,"")}</li>`).join("");
    return `<ul style='padding-left:16px;margin:4px 0'>${items}</ul>`;
  });
  html = html.replace(/\n{2,}/g,"<br/>");
  html = html.replace(/\n/g," ");
  return html;
})() }} />
                    </Box>
                  )}
                </Box>
              )}

              {/* Phase 2 */}
              {g.phase2 && g.phase2.headline && (
                <Box sx={{ mb:3 }}>
                  <Box sx={{ px:2, py:1, background:BAT_DARK, borderRadius:1, mb:1.5, display:"inline-block" }}>
                    <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase 2 — Newspaper 2027</Typography>
                  </Box>
                  <Paper elevation={1} sx={{ overflow:"hidden" }}>
                    <Box sx={{ background:`linear-gradient(135deg,${BAT_DARK},${BAT_NAVY})`, px:2.5, py:2, borderBottom:`3px solid ${BAT_GOLD}` }}>
                      <Typography variant="overline" sx={{ color:BAT_GOLD, display:"block", mb:0.5, fontSize:"0.6rem" }}>Headline</Typography>
                      <Typography variant="h6" sx={{ color:"#fff", fontWeight:800, lineHeight:1.3 }}>{g.phase2.headline}</Typography>
                    </Box>
                    <Box sx={{ p:2.5 }}>
                      <Grid container spacing={2}>
                        {["action1","action2","action3"].map((k,i) => g.phase2[k] ? (
                          <Grid item xs={12} sm={4} key={k}>
                            <Box sx={{ p:1.5, background:BAT_LIGHT, borderRadius:1, height:"100%" }}>
                              <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5, fontSize:"0.6rem" }}>Action {i+1}</Typography>
                              <Typography variant="body2" sx={{ color:"#2a3a50", lineHeight:1.6 }}>{g.phase2[k]}</Typography>
                            </Box>
                          </Grid>
                        ) : null)}
                      </Grid>
                      {g.phase2.frontlineQuote && (
                        <Box sx={{ mt:2, p:2, background:"#fffbea", borderLeft:`3px solid ${BAT_GOLD}`, borderRadius:1 }}>
                          <Typography variant="overline" sx={{ color:"#8a5900", display:"block", mb:0.5, fontSize:"0.6rem" }}>Frontline Voice</Typography>
                          <Typography variant="body2" sx={{ color:"#5a4000", lineHeight:1.6, fontStyle:"italic" }}>"{g.phase2.frontlineQuote}"</Typography>
                        </Box>
                      )}
                    </Box>
                  </Paper>
                  {g.summaries?.phase2 && (
                    <Box sx={{ mt:1.5, p:2, background:BAT_LIGHT, borderLeft:`3px solid ${BAT_NAVY}`, borderRadius:1 }}>
                      <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5, fontSize:"0.6rem" }}>AI Summary</Typography>
                      <Box sx={{
  color:"#2a3a50", fontSize:"0.82rem", lineHeight:1.6,
  "& strong":{ color:BAT_DARK, fontWeight:700 },
  "& ul":{ pl:2, mt:0.5, mb:0.5 },
  "& li":{ mb:0.3 },
  "& hr":{ border:"none", borderTop:`1px solid rgba(23,70,139,0.12)`, my:1 },
  "& blockquote":{ borderLeft:`3px solid ${BAT_GOLD}`, pl:1.5, ml:0, fontStyle:"italic", my:1 }
}} dangerouslySetInnerHTML={{ __html: (() => {
  let html = g.summaries.phase2; // change phase key here for each one
  html = html.replace(/(\|.+\|\n?)+/g,"");
  html = html.replace(/^---+$/gm,"<hr/>");
  html = html.replace(/^#{1,2} (.+)$/gm,"<strong style='display:block;font-size:0.8rem;color:#17468B;margin-top:8px'>$1</strong>");
  html = html.replace(/^#{3,4} (.+)$/gm,"<strong style='display:block;font-size:0.78rem;color:#0d2a5a;margin-top:6px'>$1</strong>");
  html = html.replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>");
  html = html.replace(/^> (.+)$/gm,"<blockquote>$1</blockquote>");
  html = html.replace(/(^[•\-\*] .+$(\n[•\-\*] .+$)*)/gm,(match)=>{
    const items = match.split("\n").filter(l=>l.trim()).map(l=>`<li>${l.replace(/^[•\-\*] /,"")}</li>`).join("");
    return `<ul style='padding-left:16px;margin:4px 0'>${items}</ul>`;
  });
  html = html.replace(/\n{2,}/g,"<br/>");
  html = html.replace(/\n/g," ");
  return html;
})() }} />
                    </Box>
                  )}
                </Box>
              )}

              {/* Phase 3 */}
              {g.phase3 && Object.values(g.phase3).some(v=>v) && (
                <Box sx={{ mb:3 }}>
                  <Box sx={{ px:2, py:1, background:BAT_NAVY, borderRadius:1, mb:1.5, display:"inline-block" }}>
                    <Typography variant="overline" sx={{ color:BAT_GOLD, fontWeight:700 }}>Phase 3 — 2027 Vision Mapping</Typography>
                  </Box>
                  <Grid container spacing={1.5}>
                    {OPM.map(el => {
                      const val = g.phase3[el.key];
                      if (!val || val.toLowerCase().includes("not addressed")) return null;
                      return (
                        <Grid item xs={12} sm={6} key={el.key}>
                          <Box sx={{ p:2, border:`1px solid rgba(250,180,30,0.3)`, borderRadius:1, height:"100%", background:"#fffbea" }}>
                            <Typography variant="overline" sx={{ color:"#8a5900", display:"block", mb:0.5, fontSize:"0.6rem" }}>{el.label}</Typography>
                            <Typography variant="body2" sx={{ color:"#2a3a50", lineHeight:1.6 }}>{val}</Typography>
                          </Box>
                        </Grid>
                      );
                    })}
                  </Grid>
                  {g.summaries?.phase3 && (
                    <Box sx={{ mt:1.5, p:2, background:BAT_LIGHT, borderLeft:`3px solid ${BAT_NAVY}`, borderRadius:1 }}>
                      <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5, fontSize:"0.6rem" }}>AI Summary</Typography>
                      <Box sx={{
  color:"#2a3a50", fontSize:"0.82rem", lineHeight:1.6,
  "& strong":{ color:BAT_DARK, fontWeight:700 },
  "& ul":{ pl:2, mt:0.5, mb:0.5 },
  "& li":{ mb:0.3 },
  "& hr":{ border:"none", borderTop:`1px solid rgba(23,70,139,0.12)`, my:1 },
  "& blockquote":{ borderLeft:`3px solid ${BAT_GOLD}`, pl:1.5, ml:0, fontStyle:"italic", my:1 }
}} dangerouslySetInnerHTML={{ __html: (() => {
  let html = g.summaries.phase3; // change phase key here for each one
  html = html.replace(/(\|.+\|\n?)+/g,"");
  html = html.replace(/^---+$/gm,"<hr/>");
  html = html.replace(/^#{1,2} (.+)$/gm,"<strong style='display:block;font-size:0.8rem;color:#17468B;margin-top:8px'>$1</strong>");
  html = html.replace(/^#{3,4} (.+)$/gm,"<strong style='display:block;font-size:0.78rem;color:#0d2a5a;margin-top:6px'>$1</strong>");
  html = html.replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>");
  html = html.replace(/^> (.+)$/gm,"<blockquote>$1</blockquote>");
  html = html.replace(/(^[•\-\*] .+$(\n[•\-\*] .+$)*)/gm,(match)=>{
    const items = match.split("\n").filter(l=>l.trim()).map(l=>`<li>${l.replace(/^[•\-\*] /,"")}</li>`).join("");
    return `<ul style='padding-left:16px;margin:4px 0'>${items}</ul>`;
  });
  html = html.replace(/\n{2,}/g,"<br/>");
  html = html.replace(/\n/g," ");
  return html;
})() }} />
                    </Box>
                  )}
                </Box>
              )}

              {/* Phase 4 */}
              {g.phase4 && Object.values(g.phase4).some(v=>v) && (
                <Box sx={{ mb:2 }}>
                  <Box sx={{ px:2, py:1, background:"#C0392B", borderRadius:1, mb:1.5, display:"inline-block" }}>
                    <Typography variant="overline" sx={{ color:"#fff", fontWeight:700 }}>Phase 4 — Roadblocks</Typography>
                  </Box>
                  <Stack spacing={1}>
                    {["r1","r2","r3"].map((k,i) => g.phase4[k] ? (
                      <Box key={k} sx={{ display:"flex", gap:2, p:2, background:"#fff", border:`1px solid rgba(192,57,43,0.2)`, borderRadius:1 }}>
                        <Box sx={{ width:32, height:32, borderRadius:1, background:"#C0392B", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                          <Typography sx={{ color:"#fff", fontWeight:900, fontSize:"0.9rem" }}>{i+1}</Typography>
                        </Box>
                        <Typography variant="body2" sx={{ color:"#2a3a50", lineHeight:1.6, pt:0.5 }}>{g.phase4[k]}</Typography>
                      </Box>
                    ) : null)}
                  </Stack>
                  {g.summaries?.phase4 && (
                    <Box sx={{ mt:1.5, p:2, background:BAT_LIGHT, borderLeft:`3px solid ${BAT_NAVY}`, borderRadius:1 }}>
                      <Typography variant="overline" sx={{ color:BAT_NAVY, display:"block", mb:0.5, fontSize:"0.6rem" }}>AI Summary</Typography>
                      <Box sx={{
  color:"#2a3a50", fontSize:"0.82rem", lineHeight:1.6,
  "& strong":{ color:BAT_DARK, fontWeight:700 },
  "& ul":{ pl:2, mt:0.5, mb:0.5 },
  "& li":{ mb:0.3 },
  "& hr":{ border:"none", borderTop:`1px solid rgba(23,70,139,0.12)`, my:1 },
  "& blockquote":{ borderLeft:`3px solid ${BAT_GOLD}`, pl:1.5, ml:0, fontStyle:"italic", my:1 }
}} dangerouslySetInnerHTML={{ __html: (() => {
  let html = g.summaries.phase4; // change phase key here for each one
  html = html.replace(/(\|.+\|\n?)+/g,"");
  html = html.replace(/^---+$/gm,"<hr/>");
  html = html.replace(/^#{1,2} (.+)$/gm,"<strong style='display:block;font-size:0.8rem;color:#17468B;margin-top:8px'>$1</strong>");
  html = html.replace(/^#{3,4} (.+)$/gm,"<strong style='display:block;font-size:0.78rem;color:#0d2a5a;margin-top:6px'>$1</strong>");
  html = html.replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>");
  html = html.replace(/^> (.+)$/gm,"<blockquote>$1</blockquote>");
  html = html.replace(/(^[•\-\*] .+$(\n[•\-\*] .+$)*)/gm,(match)=>{
    const items = match.split("\n").filter(l=>l.trim()).map(l=>`<li>${l.replace(/^[•\-\*] /,"")}</li>`).join("");
    return `<ul style='padding-left:16px;margin:4px 0'>${items}</ul>`;
  });
  html = html.replace(/\n{2,}/g,"<br/>");
  html = html.replace(/\n/g," ");
  return html;
})() }} />
                    </Box>
                  )}
                </Box>
              )}
            </Box>
          );
        })}

        <Box sx={{ textAlign:"center", pt:3, borderTop:`1px solid rgba(23,70,139,0.12)` }}><Typography variant="caption" sx={{ color:"#9bb0cc" }}>Generated by Carnelian Co · carnelianco.com · Confidential</Typography></Box></Box>
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