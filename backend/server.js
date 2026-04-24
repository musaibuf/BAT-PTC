require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { Pool } = require('pg');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
app.use(cors());
app.use(express.json());

// Set up PostgreSQL connection
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
});

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// --- DATABASE INITIALIZATION ---
// This will automatically create the tables on Render if they don't exist yet.
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
            UNIQUE(session_code, group_number)
        );
    `;

    try {
        await pool.query(createSessionsTable);
        await pool.query(createGroupsTable);
        console.log("Database tables verified/created successfully.");
    } catch (err) {
        console.error("Error initializing database tables:", err);
    }
}

// --- SESSION ROUTES ---
app.post('/api/sessions', async (req, res) => {
    const { code, numGroups, sessionName, currentPhase, startedAt } = req.body;
    try {
        await pool.query(
            'INSERT INTO sessions (code, num_groups, session_name, current_phase, started_at) VALUES ($1, $2, $3, $4, $5)',
            [code, numGroups, sessionName, currentPhase, startedAt]
        );
        for (let i = 1; i <= numGroups; i++) {
            await pool.query(
                'INSERT INTO groups (session_code, group_number) VALUES ($1, $2)',
                [code, i]
            );
        }
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
                phase3: g.phase3, phase4: g.phase4, phase3AutoMapped: g.phase3_auto_mapped, phase3Reviewed: g.phase3_reviewed
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
            phase3: g.phase3, phase4: g.phase4, phase3AutoMapped: g.phase3_auto_mapped, phase3Reviewed: g.phase3_reviewed
        });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put('/api/sessions/:code/groups/:num', async (req, res) => {
    const { joined, phase1, phase2, phase3, phase4, phase3AutoMapped, phase3Reviewed } = req.body;
    try {
        await pool.query(
            `UPDATE groups SET joined = COALESCE($1, joined), phase1 = COALESCE($2, phase1), phase2 = COALESCE($3, phase2), 
             phase3 = COALESCE($4, phase3), phase4 = COALESCE($5, phase4), phase3_auto_mapped = COALESCE($6, phase3_auto_mapped), 
             phase3_reviewed = COALESCE($7, phase3_reviewed) WHERE session_code = $8 AND group_number = $9`,
            [joined, phase1, phase2, phase3, phase4, phase3AutoMapped, phase3Reviewed, req.params.code, req.params.num]
        );
        res.json({ success: true });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- AI ROUTES ---
app.post('/api/ai/nudge', async (req, res) => {
    try {
        const msg = await anthropic.messages.create({
            model: "claude-3-5-sonnet-20241022",
            max_tokens: 200,
            system: req.body.system,
            messages: [{ role: "user", content: req.body.user }]
        });
        res.json({ text: msg.content[0].text });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/api/ai/automap', async (req, res) => {
    try {
        const msg = await anthropic.messages.create({
            model: "claude-3-5-sonnet-20241022",
            max_tokens: 700,
            system: req.body.system,
            messages: [{ role: "user", content: req.body.user }]
        });
        res.json({ text: msg.content[0].text });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- START SERVER ---
const PORT = process.env.PORT || 5000;

// Initialize the database, then start listening
initDB().then(() => {
    app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
});