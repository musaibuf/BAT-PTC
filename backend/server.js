require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { Pool } = require('pg');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
app.use(cors());
app.use(express.json());

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
});

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

async function initDB() {
    const createSessionsTable = `
        CREATE TABLE IF NOT EXISTS sessions (
            code VARCHAR(4) PRIMARY KEY,
            num_groups INT,
            session_name VARCHAR(255),
            current_phase VARCHAR(50),
            started_at BIGINT
        );
    `;
    const createGroupsTable = `
        CREATE TABLE IF NOT EXISTS groups (
            id SERIAL PRIMARY KEY,
            session_code VARCHAR(4) REFERENCES sessions(code) ON DELETE CASCADE,
            group_number INT,
            joined BOOLEAN DEFAULT FALSE,
            phase1 JSONB DEFAULT '{}',
            phase2 JSONB DEFAULT '{}',
            phase3 JSONB DEFAULT '{}',
            phase4 JSONB DEFAULT '{}',
            phase3_auto_mapped JSONB,
            phase3_reviewed BOOLEAN DEFAULT FALSE,
            summaries JSONB DEFAULT '{}',
            UNIQUE(session_code, group_number)
        );
    `;

    try {
        await pool.query(createSessionsTable);
        await pool.query(createGroupsTable);
        try { await pool.query(`ALTER TABLE groups ADD COLUMN summaries JSONB DEFAULT '{}'`); } catch(e) {}
        console.log("Database tables verified/created successfully.");
    } catch (err) { console.error("Error initializing database tables:", err); }
}

// --- SESSION ROUTES ---
app.post('/api/sessions', async (req, res) => {
    const { code, numGroups, sessionName, currentPhase, startedAt } = req.body;
    try {
        await pool.query('INSERT INTO sessions (code, num_groups, session_name, current_phase, started_at) VALUES ($1, $2, $3, $4, $5)', [code, numGroups, sessionName, currentPhase, startedAt]);
        for (let i = 1; i <= numGroups; i++) { await pool.query('INSERT INTO groups (session_code, group_number) VALUES ($1, $2)', [code, i]); }
        res.json({ success: true });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/sessions/:code', async (req, res) => {
    try {
        const result = await pool.query('SELECT * FROM sessions WHERE code = $1', [req.params.code]);
        if (result.rows.length === 0) return res.status(404).json({ error: 'Not found' });
        const s = result.rows[0];
        res.json({ code: s.code, numGroups: s.num_groups, sessionName: s.session_name, currentPhase: s.current_phase, startedAt: s.started_at });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put('/api/sessions/:code', async (req, res) => {
    try {
        await pool.query('UPDATE sessions SET current_phase = $1 WHERE code = $2', [req.body.currentPhase, req.params.code]);
        res.json({ success: true });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete('/api/sessions/:code', async (req, res) => {
    try {
        await pool.query('DELETE FROM sessions WHERE code = $1', [req.params.code]);
        res.json({ success: true });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- GROUP ROUTES ---
app.get('/api/sessions/:code/groups', async (req, res) => {
    try {
        const result = await pool.query('SELECT * FROM groups WHERE session_code = $1 ORDER BY group_number', [req.params.code]);
        const groups = {};
        result.rows.forEach(g => {
            groups[g.group_number] = {
                groupNumber: g.group_number, joined: g.joined, phase1: g.phase1, phase2: g.phase2,
                phase3: g.phase3, phase4: g.phase4, phase3AutoMapped: g.phase3_auto_mapped, phase3Reviewed: g.phase3_reviewed, summaries: g.summaries
            };
        });
        res.json(groups);
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/sessions/:code/groups/:num', async (req, res) => {
    try {
        const result = await pool.query('SELECT * FROM groups WHERE session_code = $1 AND group_number = $2', [req.params.code, req.params.num]);
        if (result.rows.length === 0) return res.status(404).json({ error: 'Not found' });
        const g = result.rows[0];
        res.json({
            groupNumber: g.group_number, joined: g.joined, phase1: g.phase1, phase2: g.phase2,
            phase3: g.phase3, phase4: g.phase4, phase3AutoMapped: g.phase3_auto_mapped, phase3Reviewed: g.phase3_reviewed, summaries: g.summaries
        });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put('/api/sessions/:code/groups/:num', async (req, res) => {
    const { joined, phase1, phase2, phase3, phase4, phase3AutoMapped, phase3Reviewed, summaries } = req.body;
    try {
        await pool.query(
            `UPDATE groups SET joined = COALESCE($1, joined), phase1 = COALESCE($2, phase1), phase2 = COALESCE($3, phase2), 
             phase3 = COALESCE($4, phase3), phase4 = COALESCE($5, phase4), phase3_auto_mapped = COALESCE($6, phase3_auto_mapped), 
             phase3_reviewed = COALESCE($7, phase3_reviewed), summaries = COALESCE($8, summaries) WHERE session_code = $9 AND group_number = $10`,
            [joined, phase1, phase2, phase3, phase4, phase3AutoMapped, phase3Reviewed, summaries, req.params.code, req.params.num]
        );
        res.json({ success: true });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- AI ROUTES ---
app.post('/api/ai/nudge', async (req, res) => {
    try {
        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 3000,
            system: req.body.system,
            messages: [{ role: "user", content: req.body.user }]
        });
        res.json({ text: msg.content[0].text });
    } catch (err) {
        console.error("Nudge Error:", err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/ai/automap', async (req, res) => {
    try {
        let userContent = req.body.user;
        if (typeof userContent === 'object') userContent = JSON.stringify(userContent);

        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 3000,
            system: req.body.system,
            messages: [{ role: "user", content: userContent }]
        });
        
        let cleanText = msg.content[0].text.replace(/```json\n?/gi, '').replace(/```\n?/g, '').trim();
        res.json({ text: cleanText });
    } catch (err) {
        console.error("Automap Error:", err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/ai/summarize', async (req, res) => {
    try {
        const { data, phase } = req.body;

        if (!data || typeof data !== 'string' || data.trim() === "" || data.trim() === '"{}"' || data.trim() === '{}') {
            return res.json({ text: "Not enough data to summarize yet." });
        }

        try {
            const parsed = JSON.parse(data);
            if (typeof parsed === 'object' && Object.keys(parsed).length === 0) {
                return res.json({ text: "Not enough data to summarize yet." });
            }
        } catch(e) { /* plain text, fine */ }

        let sys = "";

        if (phase === "phase1") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.
You are analyzing GROUP INPUT from Phase 1: Current State Assessment.

Return your analysis using ONLY this exact plain-text format — no markdown headers, no asterisks, no bold, no hash symbols:

SUMMARY
Write one punchy paragraph (3-4 sentences) capturing the overall culture and operational reality. Be direct and honest.

COMMON THEMES
- [Theme name]: one sentence on what this reveals about culture or behaviour
- [Theme name]: one sentence on what this reveals about culture or behaviour
- [Theme name]: one sentence on what this reveals about culture or behaviour

PROBLEMATIC BEHAVIOURS
- [Behaviour]: why this is a blocker or cultural risk
- [Behaviour]: why this is a blocker or cultural risk
- [Behaviour]: why this is a blocker or cultural risk

ACTIONABLE STEPS
- [Step]: what must change and why
- [Step]: what must change and why
- [Step]: what must change and why

Be specific to what the group actually said. No corporate fluff.`;

        } else if (phase === "phase2") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.
You are analyzing GROUP INPUT from Phase 2: Newspaper Headline Exercise.

Return your analysis using ONLY this exact plain-text format — no markdown headers, no asterisks, no bold, no hash symbols:

SUMMARY
State the headline this group produced, then 2-3 sentences on what it reveals about what they want to change culturally and behaviourally.

WHAT THIS GROUP IS ASKING FOR
2-3 sentences on the frustrations, hopes, and cultural shifts this group believes are necessary. Read between the lines.

ACTIONABLE STEPS
- [Step]: what behaviour or system needs to change to make this headline real
- [Step]: what needs to change
- [Step]: what needs to change
- [Step]: what needs to change

CULTURAL SHIFTS REQUIRED
- From [current state] to [desired state]: why this shift matters
- From [current state] to [desired state]: why this shift matters
- From [current state] to [desired state]: why this shift matters

Be sharp, specific, and connect back to Phase 1 problems.`;

        } else if (phase === "phase3") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.
You are analyzing GROUP INPUT from Phase 3: Gap Mapping.

Return your analysis using ONLY this exact plain-text format — no markdown headers, no asterisks, no bold, no hash symbols:

SUMMARY
One sharp paragraph (3-4 sentences) on the overall size and nature of the transformation gap. Be honest about how significant it is.

KEY GAPS
- [Design element]: current state vs desired state and what this gap means culturally
- [Design element]: current state vs desired state and what this gap means culturally
- [Design element]: current state vs desired state and what this gap means culturally

BEHAVIOURAL SHIFTS TO CLOSE THE GAP
- [Shift]: specific behaviour that must change and what it unlocks
- [Shift]: specific behaviour that must change and what it unlocks
- [Shift]: specific behaviour that must change and what it unlocks

ACTIONABLE STEPS
- [Step]: concrete action and why it closes a specific gap
- [Step]: concrete action and why it closes a specific gap
- [Step]: concrete action and why it closes a specific gap

BIGGEST RISK IF NOT CLOSED
One direct sentence on what happens to the culture if this gap remains unaddressed.`;

        } else if (phase === "phase4") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.
You are analyzing GROUP INPUT from Phase 4: Roadblocks.

Return your analysis using ONLY this exact plain-text format — no markdown headers, no asterisks, no bold, no hash symbols:

SUMMARY
One sharp paragraph on the most critical blockers. Distinguish between systemic blockers and behavioural or cultural blockers.

TOP ROADBLOCKS
- [Roadblock]: systemic or behavioural — why this is the real blocker
- [Roadblock]: systemic or behavioural — why this is the real blocker
- [Roadblock]: systemic or behavioural — why this is the real blocker

ACTIONABLE STEPS
- [Step]: what leadership must do to remove this blocker
- [Step]: what leadership must do to remove this blocker
- [Step]: what leadership must do to remove this blocker

TRANSFORMATION STORY
3-4 sentences weaving together the full arc: where they are today, what they envision, how big the gap is, and what stands in the way. Read like the opening of a transformation brief.`;

        } else {
            sys = `You are an organizational culture analyst. Summarize the provided group workshop input into clear insights about culture, behaviour, and transformation. Use plain bullet points only — no markdown headers, no asterisks, no bold formatting.`;
        }

        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 3000,
            system: sys,
            messages: [{ role: "user", content: data }]
        });

        res.json({ text: msg.content[0].text });
    } catch (err) {
        console.error("Summarize Error from Anthropic:", err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/ai/themes', async (req, res) => {
    try {
        const sys = `You are a senior organizational culture strategist producing the final synthesis report from a full-day Culture Gap Assessment workshop.

You have been given the complete data from ALL groups across ALL phases. Your job is to produce a deep, honest, and actionable analysis.

Return ONLY valid JSON. No markdown, no backticks, no preamble. Use this exact structure:
{
  "executiveSummary": "4-5 sentences covering: todays cultural reality, the shared vision, the transformation gap, common blockers, and what leadership must prioritize. Be direct and specific.",

  "themes": [
    "Trust Erosion — one sentence on what this reveals about culture and behaviour",
    "Structural Fragmentation — one sentence on what this reveals",
    "Human-Performance Disconnect — one sentence on what this reveals"
  ],

  "behaviouralShifts": [
    "From hoarding decisions at the top → To distributing authority with clear accountability: without this shift no other change will stick",
    "From silence out of fear → To candour as a protected behaviour: psychological safety is the prerequisite for every other intervention",
    "From siloed execution → To cross-functional ownership: the commercial vision requires structural permission to collaborate"
  ],

  "priorityActions": [
    "Leadership Behaviour Reset — senior leaders must visibly model intellectual humility and invite challenge before any programme will land",
    "Structural Silo Intervention — establish cross-functional teams with shared goals and metrics so collaboration has structural permission",
    "Recognition and Safety System — build merit-based recognition and protected channels for upward feedback as the long-term cultural anchor"
  ],

  "groupInsights": {
    "common": "Write 2-3 sentences describing ONLY what every group explicitly agreed on — the problems, themes, or ambitions that appeared consistently across ALL groups without exception.",
    "divergence": "Write 2-3 sentences describing ONLY where groups had DIFFERENT views, priorities, or levels of readiness — specific contrasts between groups. If only one group participated, describe what was absent from other groups and what that silence may signal about engagement or fear."
  },

  "seniorManagement": {
    "context": "2-3 sentences grounding these commitments in the specific cultural reality this workshop surfaced — name the exact dynamics senior leaders are directly accountable for creating or allowing, and why the transformation cannot proceed without their visible, behavioural change first.",
    "themes": [
      "Theme Name — 3-4 sentences. Explain what this theme means specifically at the senior leadership level: what decision-making patterns, communication habits, or structural choices are sustaining this theme. Name what senior leaders are doing or not doing that keeps this theme alive. Be direct about accountability.",
      "Theme Name — 3-4 sentences on the same depth and structure.",
      "Theme Name — 3-4 sentences on the same depth and structure."
    ],
    "behaviours": [
      "START: Name the behaviour, then 2-3 sentences explaining exactly what it looks like in practice, why it is currently absent, and what cultural signal it sends when senior leaders do it consistently. Be specific enough that a leader knows precisely what to do differently in their next meeting or decision.",
      "STOP: Name the behaviour, then 2-3 sentences explaining what this behaviour currently looks like, the cultural damage it causes, and what changes when it stops. Connect it directly to something groups raised in the workshop.",
      "SUSTAIN: Name the behaviour, then 2-3 sentences explaining why this must be protected and reinforced, what it would cost the transformation if it were lost, and how senior leaders should actively signal they are sustaining it."
    ],
    "actions": [
      "Action Title — Specific: describe exactly what will be done, by whom, and in what format or forum. Measurable: state the concrete metric or observable output that proves this happened — not a feeling, a fact. Attainable: explain why this is realistic given current capacity and structure. Relevant: name the exact cultural gap from the workshop this closes and why it matters for the transformation. Time-bound: give a specific deadline — not a range, a date or number of weeks from the workshop.",
      "Action Title — same full SMART structure, different gap addressed.",
      "Action Title — same full SMART structure, different gap addressed."
    ]
  },

  "middleManagement": {
    "context": "2-3 sentences grounding these commitments in the position middle managers occupy — translating leadership intent into daily team reality while absorbing pressure from both directions. Name what the workshop data revealed about where this translation is currently breaking down and what that costs the organisation.",
    "themes": [
      "Theme Name — 3-4 sentences. Explain what this theme looks like at the team and department level: what patterns middle managers are reinforcing or failing to interrupt. Be specific about the daily behaviours and micro-decisions that sustain this theme at their level of the organisation.",
      "Theme Name — 3-4 sentences on the same depth and structure.",
      "Theme Name — 3-4 sentences on the same depth and structure."
    ],
    "behaviours": [
      "START: Name the behaviour, then 2-3 sentences explaining exactly what it looks like in a team setting, why it is currently missing, and what it unlocks for team trust, performance, or clarity when middle managers adopt it consistently.",
      "STOP: Name the behaviour, then 2-3 sentences explaining how this behaviour currently manifests at team level, the specific damage it does to psychological safety or performance, and what becomes possible when it stops.",
      "SUSTAIN: Name the behaviour, then 2-3 sentences explaining why this must be protected at team level, what erosion looks like under pressure, and how managers should actively signal they are sustaining it even when it is difficult."
    ],
    "actions": [
      "Action Title — Specific: describe exactly what the manager will do, with their team or peers, in what cadence or format. Measurable: state the concrete metric or output — a number, a deliverable, a visible change. Attainable: explain why a manager with a full workload can realistically execute this. Relevant: name the exact gap this closes at team level. Time-bound: specific deadline in weeks from the workshop.",
      "Action Title — same full SMART structure, different gap addressed.",
      "Action Title — same full SMART structure, different gap addressed."
    ]
  },

  "individualContributors": {
    "context": "2-3 sentences grounding these commitments in what individual contributors actually said — their specific frustrations, the things they felt unable to say or do, and what they need in terms of permission, safety, or clarity to show up differently. Make this feel like it was written for them, not about them.",
    "themes": [
      "Theme Name — 3-4 sentences. Explain what this theme means from the ground level: how it shows up in day-to-day work, what it costs individuals personally and professionally, and what shifts in their immediate environment would need to happen for this theme to change. Acknowledge the structural constraints they operate within.",
      "Theme Name — 3-4 sentences on the same depth and structure.",
      "Theme Name — 3-4 sentences on the same depth and structure."
    ],
    "behaviours": [
      "START: Name the behaviour, then 2-3 sentences explaining exactly what it looks like in daily work, why it has felt unsafe or unnecessary until now, and what changes for the individual and their team when they begin doing it consistently.",
      "STOP: Name the behaviour, then 2-3 sentences explaining how this behaviour currently shows up as a coping mechanism or habit, why it made sense in the old culture, and what it costs the team or the individual when it continues.",
      "SUSTAIN: Name the behaviour, then 2-3 sentences explaining why this individual habit or practice is a genuine asset to the transformation, what erodes it under pressure, and how individuals can protect it even when the system around them is slow to change."
    ],
    "actions": [
      "Action Title — Specific: describe exactly what the individual will do, in what context, and with what frequency. Measurable: state how the individual or their manager will know this is happening — a habit tracked, a conversation logged, a piece of work produced. Attainable: confirm this is within the individual's direct control and does not require structural permission they do not currently have. Relevant: connect this directly to a frustration or gap the groups named in the workshop. Time-bound: specific deadline in days or weeks — short enough to feel immediate.",
      "Action Title — same full SMART structure, different gap addressed.",
      "Action Title — same full SMART structure, different gap addressed."
    ]
  }
}

CRITICAL RULES:
- The common and divergence fields MUST contain different content. Never repeat the same text in both.
- divergence must describe genuine differences or contrasts between groups, not similarities.
- All string values must be plain prose — no markdown, no bullet symbols, no asterisks inside the JSON values.
- Return ONLY the JSON object.`;

        const msg = await anthropic.messages.create({
            model: "claude-opus-4-7",
            max_tokens: 4000, // Increased to 4000 to handle the larger JSON output
            system: sys,
            messages: [{ role: "user", content: req.body.data }]
        });

        // Strip markdown to prevent JSON parse errors on frontend
        let cleanJsonText = msg.content[0].text.replace(/```json\n?/gi, '').replace(/```\n?/g, '').trim();
        res.json({ text: cleanJsonText });
    } catch (err) {
        console.error("Themes Error:", err);
        res.status(500).json({ error: err.message });
    }
});

const PORT = process.env.PORT || 5000;
initDB().then(() => { app.listen(PORT, () => console.log(`Server running on port ${PORT}`)); });