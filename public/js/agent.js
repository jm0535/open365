/*
 * Open365 Agent — a Copilot-style side panel for every editor. The user
 * brings their own API key (stored in localStorage, sent only to the chosen
 * provider). The agent receives a compact snapshot of the open document plus
 * a list of editor tools, and replies with JSON actions that are executed
 * live — each edit lands through the editor's own mutation path, so undo,
 * autosave and presence behave exactly as if a person made the change.
 *
 * Providers: any OpenAI-compatible /chat/completions endpoint (OpenAI,
 * OpenRouter, Ollama, LM Studio, vLLM…) or Anthropic's /v1/messages.
 */
import { h, toast, toastErr, modal } from './util.js';
import { icon } from './icons.js';

const CFG_KEY = 'o365.agent.cfg';
const ROUNDS_MAX = 8;          // agent loop iterations per task
const STEP_MS = 320;           // pause between actions so the user can watch

export function agentConfig() {
  try { return JSON.parse(localStorage.getItem(CFG_KEY)) || {}; } catch { return {}; }
}
export function setAgentConfig(cfg) {
  try { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); } catch { /* private mode */ }
}
export function hasAgentKey() { return !!agentConfig().key; }

/* ------------------------------------------------------------ transport */
async function chat(messages, cfg) {
  if (cfg.provider === 'anthropic') {
    const res = await fetch((cfg.base || 'https://api.anthropic.com').replace(/\/+$/, '') + '/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': cfg.key,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify({
        model: cfg.model || 'claude-sonnet-4-6',
        max_tokens: 4096,
        system: messages.find(m => m.role === 'system')?.content || '',
        messages: messages.filter(m => m.role !== 'system')
      })
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error?.message || ('Anthropic error ' + res.status));
    return (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n');
  }
  const res = await fetch((cfg.base || 'https://api.openai.com/v1').replace(/\/+$/, '') + '/chat/completions', {
    method: 'POST',
    headers: { 'content-type': 'application/json', authorization: 'Bearer ' + cfg.key },
    body: JSON.stringify({ model: cfg.model || 'gpt-4o-mini', messages, temperature: 0.2 })
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error?.message || ('Provider error ' + res.status));
  return data.choices?.[0]?.message?.content || '';
}

/* ------------------------------------------------------------- parsing */
/* The model answers with JSON — possibly wrapped in a fenced code block. */
function parseTurn(text) {
  const clean = String(text || '').replace(/```(?:json)?/gi, '');
  const start = clean.indexOf('{');
  const end = clean.lastIndexOf('}');
  if (start < 0 || end <= start) return { say: clean.trim(), actions: [], done: true };
  try {
    const obj = JSON.parse(clean.slice(start, end + 1));
    const actions = Array.isArray(obj.actions) ? obj.actions
      : (obj.tool ? [obj] : []);
    return {
      say: String(obj.say || obj.message || ''),
      actions: actions.filter(a => a && typeof a.tool === 'string'),
      done: obj.done !== false && obj.continue !== true
    };
  } catch {
    return { say: clean.trim(), actions: [], done: true };
  }
}

/* ------------------------------------------------------------- panel UI */
let openPanel = null;

export function agentSettingsModal() {
  const cfg = agentConfig();
  const prov = h('select', { class: 'sel', style: { width: '100%' } },
    h('option', { value: 'openai', text: 'OpenAI-compatible endpoint', selected: cfg.provider !== 'anthropic' ? '' : null }),
    h('option', { value: 'anthropic', text: 'Anthropic', selected: cfg.provider === 'anthropic' ? '' : null }));
  const base = h('input', { value: cfg.base || '', placeholder: 'https://api.openai.com/v1', style: { width: '100%' } });
  const key = h('input', { type: 'password', value: cfg.key || '', placeholder: 'sk-…', style: { width: '100%' } });
  const model = h('input', { value: cfg.model || '', placeholder: 'gpt-4o-mini · claude-sonnet-4-6 · llama3.1', style: { width: '100%' } });
  const box = modal({
    title: 'Agent settings',
    body: h('div', {},
      h('p', { class: 'small muted', text: 'Your key stays in this browser (localStorage) and is sent only to the provider you configure. Nothing goes through Open365.' }),
      h('div', { class: 'field' }, h('label', { text: 'Provider' }), prov),
      h('div', { class: 'field' }, h('label', { text: 'Base URL (optional)' }), base),
      h('div', { class: 'field' }, h('label', { text: 'API key' }), key),
      h('div', { class: 'field' }, h('label', { text: 'Model' }), model)),
    footer: [
      h('button', { class: 'btn', text: 'Cancel', onclick: () => box.close() }),
      h('button', { class: 'btn primary', text: 'Save', onclick: () => {
        setAgentConfig({ provider: prov.value, base: base.value.trim(), key: key.value.trim(), model: model.value.trim() });
        box.close();
        toast(cfg.key && key.value.trim() !== cfg.key ? 'Agent key updated' : 'Agent settings saved', 'ok');
      } })
    ]
  });
}

function systemPrompt(appName, tools) {
  const lines = tools.map(t => '- ' + t.name + (t.args ? '(' + t.args + ')' : '') + ': ' + t.desc).join('\n');
  return [
    'You are an autonomous agent inside ' + appName + ' (an Office-style app), editing the open document directly.',
    'You act through these tools only:',
    lines,
    '',
    'Reply with ONE JSON object only, no prose around it:',
    '{"say":"short note to the user","actions":[{"tool":"name","args":{...}}],"done":true}',
    '',
    'Rules:',
    '- "actions" is the ordered list of tool calls for THIS step. Keep it under 25 calls.',
    '- Set "done":false when you need another step after seeing the results; the document snapshot and action results are sent back to you.',
    '- Use read tools before writing when unsure of the current content.',
    '- Never invent data the user did not ask for. Confirm destructive moves in "say".',
    '- When the task is finished set "done":true and "actions":[].'
  ].join('\n');
}

const wait = (ms) => new Promise(r => setTimeout(r, ms));

/**
 * Mount an agent panel bound to an editor.
 * ctx = { app: 'Sheets', title: 'Q3 budget', snapshot: () => string,
 *         tools: [{name, desc, args, call(args) -> string}], flash?(el) }
 */
export function openAgentPanel(ctx) {
  if (openPanel) { openPanel.close(); openPanel = null; }
  const cfg = agentConfig();
  const log = h('div', { class: 'agentlog' });
  const input = h('textarea', { class: 'agentin', placeholder: 'Tell the agent what to do… (e.g. "add a totals row")', rows: 2 });
  const sendBtn = h('button', { class: 'btn primary', html: icon('send', 15) + '<span>Run</span>' });
  const stopBtn = h('button', { class: 'btn', text: 'Stop', style: { display: 'none' } });
  const statusEl = h('div', { class: 'agentstatus small muted' });
  const badge = h('span', { class: 'pill accent agentbadge', style: { display: 'none' }, text: 'Agent working…' });
  document.querySelector('.o365-title')?.append(badge);

  let busy = false, stopped = false;
  const messages = [{ role: 'system', content: systemPrompt(ctx.app, ctx.tools) }];

  const say = (who, text) => {
    log.append(h('div', { class: 'agentmsg ' + who }, h('b', { text: who === 'me' ? 'You' : who === 'act' ? '' : 'Agent' }), h('span', { text })));
    log.scrollTop = log.scrollHeight;
  };
  const note = (text) => {
    log.append(h('div', { class: 'agentnote small muted', text }));
    log.scrollTop = log.scrollHeight;
  };

  const tool = (name) => ctx.tools.find(t => t.name === name);

  async function runActions(turn) {
    const results = [];
    for (const a of turn.actions.slice(0, 40)) {
      if (stopped) break;
      const t = tool(a.tool);
      statusEl.textContent = '→ ' + a.tool;
      if (!t) { results.push({ tool: a.tool, ok: false, error: 'no such tool' }); note('✗ ' + a.tool + ' — unknown'); continue; }
      try {
        const out = await t.call(a.args || {});
        results.push({ tool: a.tool, ok: true, result: String(out ?? 'ok').slice(0, 2000) });
        note('→ ' + a.tool + ' ' + JSON.stringify(a.args || {}).slice(0, 90) + ' ⇒ ' + String(out ?? 'ok').slice(0, 120));
      } catch (err) {
        results.push({ tool: a.tool, ok: false, error: String(err.message || err).slice(0, 300) });
        note('✗ ' + a.tool + ' failed: ' + (err.message || err));
      }
      await wait(STEP_MS);
    }
    return results;
  }

  async function run(task) {
    const c = agentConfig();
    if (!c.key) { agentSettingsModal(); return; }
    busy = true; stopped = false;
    sendBtn.style.display = 'none'; stopBtn.style.display = '';
    badge.style.display = ''; statusEl.textContent = 'Thinking…';
    messages.push({ role: 'user', content: task });
    try {
      for (let round = 0; round < ROUNDS_MAX && !stopped; round++) {
        const reply = parseTurn(await chat(messages.concat({
          role: 'user',
          content: 'Current document state:\n' + ctx.snapshot() + (round ? '\n\nResults of your last actions were sent above. Continue or finish.' : '')
        }), c));
        if (reply.say) say('agent', reply.say);
        messages.push({ role: 'assistant', content: reply.say ? JSON.stringify({ say: reply.say }) : '{}' });
        if (!reply.actions.length || reply.done && !reply.actions.length) {
          if (round === 0 && !reply.actions.length && !reply.say) say('agent', 'Done.');
          break;
        }
        const results = await runActions(reply);
        if (stopped) break;
        if (reply.done !== false) { /* model asked for another look anyway */ }
        if (reply.done) { /* executed final batch */ break; }
        messages.push({ role: 'user', content: 'Action results:\n' + JSON.stringify(results) });
      }
      if (stopped) note('Stopped by user.');
    } catch (err) {
      say('agent', '');
      note('Error: ' + (err.message || err));
      toastErr(err.message || 'Agent failed');
    }
    busy = false; stopped = false;
    sendBtn.style.display = ''; stopBtn.style.display = 'none';
    badge.style.display = 'none'; statusEl.textContent = '';
  }

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      const t = input.value.trim();
      if (!t || busy) return;
      input.value = '';
      say('me', t);
      run(t);
    }
    e.stopPropagation();
  });
  sendBtn.addEventListener('click', () => {
    const t = input.value.trim();
    if (!t || busy) return;
    input.value = '';
    say('me', t);
    run(t);
  });
  stopBtn.addEventListener('click', () => { stopped = true; });

  const panel = h('aside', { class: 'agentpane' },
    h('div', { class: 'agenthead' },
      h('span', { class: 'agentmark', html: icon('bot', 15) }),
      h('b', { text: ctx.app + ' agent' }),
      h('span', { class: 'grow' }),
      h('button', { class: 'btn-icon', title: 'Agent settings', html: icon('settings', 14), onclick: agentSettingsModal }),
      h('button', { class: 'btn-icon', title: 'Close', html: icon('close', 14), onclick: () => close() })),
    cfg.key ? null : h('button', { class: 'agentkeys', text: 'Add an API key to enable the agent', onclick: agentSettingsModal }),
    log,
    statusEl,
    h('div', { class: 'agentbar' }, input, h('div', { class: 'center-row' }, sendBtn, stopBtn))
  );

  function close() { panel.remove(); badge.remove(); if (openPanel === api) openPanel = null; }
  const api = { el: panel, close, say, note, run };
  openPanel = api;
  return api;
}

/* ---------------------------------------------------------------- helper */
/* Briefly highlight an element the agent just touched. */
export function agentFlash(el) {
  if (!el || !el.classList) return;
  el.classList.add('agentflash');
  setTimeout(() => el.classList.remove('agentflash'), 900);
}
