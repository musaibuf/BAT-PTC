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
// --- AI ROUTES ---
app.post('/api/ai/nudge', async (req, res) => {
    try {
        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 300,
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
        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 1500,
            system: req.body.system,
            messages: [{ role: "user", content: req.body.user }]
        });
        res.json({ text: msg.content[0].text });
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

Your job is to extract and surface the REAL cultural and behavioural picture from what this group has shared.

Return your analysis in this exact format:

**Current State Summary**
One punchy paragraph (3-4 sentences) capturing the overall culture and operational reality this group is living in right now. Be direct and honest — no corporate fluff.

**Common Themes (3-4)**
- [Theme 1]: One sentence explaining what this theme reveals about the culture or behaviour
- [Theme 2]: One sentence explaining what this theme reveals about the culture or behaviour
- [Theme 3]: One sentence explaining what this theme reveals about the culture or behaviour
- [Theme 4 if applicable]: One sentence

**Problematic Behaviours Identified**
- [Behaviour]: Why this is a blocker or cultural risk
- [Behaviour]: Why this is a blocker or cultural risk
- [Behaviour]: Why this is a blocker or cultural risk

Be specific to what the group actually said. Do not generalize vaguely.`;

        } else if (phase === "phase2") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.

You are analyzing GROUP INPUT from Phase 2: Newspaper Headline Exercise — where participants imagined a future headline about their organization's transformation.

Your job is to extract ambition, cultural signals, and the actionable steps implied in what this group envisions.

Return your analysis in this exact format:

**Headline Insight**
State the most powerful or representative headline this group produced, then in 2-3 sentences explain what it reveals about what this group truly wants to change culturally and behaviourally.

**What This Group Is Really Asking For**
2-3 sentences going deeper — what does this vision tell us about the frustrations, hopes, and cultural shifts this group believes are necessary? Read between the lines.

**Implied Actionable Steps**
- [Step]: What behaviour or system needs to change to make this headline real
- [Step]: What behaviour or system needs to change
- [Step]: What behaviour or system needs to change
- [Step if applicable]

**Cultural & Behaviour Shift Required**
- [Shift]: From [current state] → To [desired state]
- [Shift]: From [current state] → To [desired state]
- [Shift]: From [current state] → To [desired state]

Be sharp, specific, and connect directly back to what Phase 1 likely revealed about current problems.`;

        } else if (phase === "phase3") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.

You are analyzing GROUP INPUT from Phase 3: Gap Mapping — where participants mapped the gap between today's reality and the 2027 vision across key design elements.

Your job is to synthesize the gaps into a clear picture of what needs to change, connecting it back to the cultural and behavioural themes from Phase 1 and Phase 2.

Return your analysis in this exact format:

**Gap Summary**
One sharp paragraph (3-4 sentences) describing the overall size and nature of the transformation gap this group has identified. Be honest about how significant the distance is.

**Key Gaps by Design Element**
- [Element]: Current state vs desired state — what this gap means culturally
- [Element]: Current state vs desired state — what this gap means culturally
- [Element]: Current state vs desired state — what this gap means culturally

**Behavioural Shifts Needed to Close the Gap**
- [Shift]: Specific behaviour that must change and what it unlocks
- [Shift]: Specific behaviour that must change and what it unlocks
- [Shift]: Specific behaviour that must change and what it unlocks

**Connection to Earlier Phases**
2-3 sentences connecting these gaps back to the problematic behaviours from Phase 1 and the vision from Phase 2. Show the through-line of the transformation story.

**Biggest Risk If Gap Is Not Closed**
One direct sentence about what happens to the culture if this gap remains unaddressed.`;

        } else if (phase === "phase4") {
            sys = `You are an organizational culture analyst facilitating a corporate transformation workshop.

You are analyzing GROUP INPUT from Phase 4: Roadblocks & Commitments — where participants identified what stands in the way and what they personally commit to changing.

Your job is to synthesize the blockers and commitments into an honest, actionable picture that connects the full arc from Phase 1 through Phase 4.

Return your analysis in this exact format:

**Roadblock Summary**
One sharp paragraph identifying the most critical blockers this group named. Distinguish between systemic blockers (structures, processes) and behavioural/cultural blockers (mindsets, habits, norms).

**Top Roadblocks**
- [Roadblock]: Type (systemic / behavioural) — why this is the real blocker
- [Roadblock]: Type (systemic / behavioural) — why this is the real blocker
- [Roadblock]: Type (systemic / behavioural) — why this is the real blocker

**Commitments Made**
- [Commitment]: What this signals about this group's readiness to change
- [Commitment]: What this signals about this group's readiness to change
- [Commitment]: What this signals about this group's readiness to change

**Full Arc — This Group's Transformation Story**
3-4 sentences that weave together the journey: where they are today (Phase 1), what they envision (Phase 2), how big the gap is (Phase 3), and what stands in the way and what they're committing to (Phase 4). This should read like the opening paragraph of a transformation brief.

**Recommended Priority Actions**
- [Action 1]: Why this must happen first
- [Action 2]: Why this follows
- [Action 3]: Why this is the long-term cultural anchor`;

        } else {
            // fallback generic
            sys = `You are an organizational culture analyst. Summarize the provided group workshop input into clear insights about culture, behaviour, and transformation. Be specific, direct, and actionable. Use bullet points where appropriate.`;
        }

        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 800,
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

Return ONLY valid JSON with the following structure:
{
  "executiveSummary": "A 4-5 sentence executive summary that covers: (1) the overall cultural reality today, (2) the shared vision across groups, (3) the size and nature of the transformation gap, (4) the common blockers, and (5) what this organization must prioritize to make the shift real. Be direct — this is for leadership.",

  "themes": [
    "Theme 1: [Name] — [One sentence on what this theme reveals about culture and behaviour across groups]",
    "Theme 2: [Name] — [One sentence on what this theme reveals about culture and behaviour across groups]",
    "Theme 3: [Name] — [One sentence on what this theme reveals about culture and behaviour across groups]"
  ],

  "behaviouralShifts": [
    "From [current behaviour] → To [required behaviour]: [Why this shift is critical]",
    "From [current behaviour] → To [required behaviour]: [Why this shift is critical]",
    "From [current behaviour] → To [required behaviour]: [Why this shift is critical]"
  ],

  "priorityActions": [
    "Action 1: [What] — [Why this must happen first across the organization]",
    "Action 2: [What] — [Why this follows]",
    "Action 3: [What] — [Why this is the long-term cultural anchor]"
  ],

  "groupInsights": {
    "common": "2-3 sentences on what ALL groups agreed on — the universal truths from today's session.",
    "divergence": "2-3 sentences on where groups differed significantly — what this tells leadership about pockets of different readiness or different realities."
  }
}

Return ONLY the JSON object. No markdown, no backticks, no preamble.`;

        const msg = await anthropic.messages.create({
            model: "claude-sonnet-4-6",
            max_tokens: 3000,
            system: sys,
            messages: [{ role: "user", content: req.body.data }]
        });

        res.json({ text: msg.content[0].text });
    } catch (err) {
        console.error("Themes Error:", err);
        res.status(500).json({ error: err.message });
    }
});

const PORT = process.env.PORT || 5000;
initDB().then(() => { app.listen(PORT, () => console.log(`Server running on port ${PORT}`)); });