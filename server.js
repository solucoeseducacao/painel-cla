'use strict';
const express      = require('express');
const cookieParser = require('cookie-parser');
const multer       = require('multer');
const https        = require('https');
const crypto       = require('crypto');
const path         = require('path');
const Anthropic    = require('@anthropic-ai/sdk');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType } = require('docx');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT   = process.env.PORT || 3000;

// ── Helpers de ambiente ────────────────────────────────────────────────────
function requireEnv(name) {
  const val = process.env[name];
  if (!val) throw new Error(`Variável de ambiente obrigatória não configurada: ${name}`);
  return val;
}

// ── Auth ───────────────────────────────────────────────────────────────────
function sign(data) {
  return crypto.createHmac('sha256', requireEnv('SESSION_SECRET')).update(data).digest('base64url');
}
function makeToken(email) {
  const payload = Buffer.from(JSON.stringify({ email, exp: Date.now() + 86400000 * 7 })).toString('base64url');
  return `${payload}.${sign(payload)}`;
}
function verifyToken(token) {
  if (!token) return null;
  const dotIdx = token.indexOf('.');
  if (dotIdx < 0) return null;
  const payload = token.slice(0, dotIdx);
  const sig     = token.slice(dotIdx + 1);
  if (!payload || !sig) return null;
  try {
    const expected    = sign(payload);
    const sigBuf      = Buffer.from(sig,      'base64url');
    const expectedBuf = Buffer.from(expected, 'base64url');
    if (sigBuf.length !== expectedBuf.length) return null;
    if (!crypto.timingSafeEqual(sigBuf, expectedBuf)) return null;
    const data = JSON.parse(Buffer.from(payload, 'base64url').toString());
    if (data.exp < Date.now()) return null;
    if (data.email !== requireEnv('ALLOWED_EMAIL')) return null;
    return data;
  } catch { return null; }
}
function requireSession(req, res, next) {
  const token = req.cookies?.cla_session;
  const data  = verifyToken(token);
  if (!data) return res.status(401).json({ error: 'Sessão inválida ou expirada.' });
  req.user = data;
  next();
}

// ── HTTP helpers ───────────────────────────────────────────────────────────
function httpsPost(url, data) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams(data).toString();
    const u    = new URL(url);
    const req  = https.request({
      hostname: u.hostname, path: u.pathname, method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}
function httpsGet(url, token) {
  return new Promise((resolve, reject) => {
    const u   = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: `Bearer ${token}` }
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    req.end();
  });
}

// ── Anthropic ──────────────────────────────────────────────────────────────
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const MODEL  = 'claude-sonnet-4-6';

// ── BUDGET ─────────────────────────────────────────────────────────────────
const BUDGET = {
  A: {
    '2': { conv: 20, quiz: 10, entrada: 10, grp: 20, reg: 15, cpd: 15 },
    '3': { conv: 30, quiz: 10, entrada: 10, grp: 35, reg: 20, cpd: 15 },
    '4': { conv: 45, quiz: 10, entrada: 15, grp: 50, reg: 25, cpd: 15 }
  },
  B: {
    '2': { leit: 30, quiz: 10, grp: 20, reg: 15, cpd: 15 },
    '3': { leit: 45, quiz: 10, grp: 35, reg: 15, cpd: 15 },
    '4': { leit: 65, quiz: 10, grp: 50, reg: 20, cpd: 15 }
  },
  C: {
    '2': { conv: 15, lit: 15, quiz: 10, grp: 20, reg: 15, cpd: 15 },
    '3': { conv: 20, lit: 20, quiz: 10, grp: 35, reg: 20, cpd: 15 },
    '4': { conv: 30, lit: 30, quiz: 10, grp: 45, reg: 25, cpd: 20 }
  }
};

// ── docx helpers ───────────────────────────────────────────────────────────
const A4_PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } };
const BDR     = { style: BorderStyle.SINGLE, size: 1, color: '999999' };
const BORDERS = { top: BDR, bottom: BDR, left: BDR, right: BDR };

function tc(children, opts = {}) {
  return new TableCell({
    borders: BORDERS,
    shading: { fill: opts.fill || 'FFFFFF', type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width:   { size: opts.w || 4513, type: WidthType.DXA },
    children
  });
}
function makeHeader(disc, semana, docName) {
  const W1 = 5416, W2 = 3610;
  return new Table({
    width: { size: 9026, type: WidthType.DXA }, columnWidths: [W1, W2],
    rows: [
      new TableRow({ children: [
        tc([
          new Paragraph({ children: [new TextRun({ text: 'INSTITUTO FEDERAL FLUMINENSE', bold: true, size: 18, font: 'Arial' })] }),
          new Paragraph({ children: [new TextRun({ text: 'Campus Campos Centro · Licenciatura em Letras — Português e Literaturas', size: 17, font: 'Arial' })] }),
          new Paragraph({ children: [new TextRun({ text: `Disciplina: ${disc || ''}`, size: 17, font: 'Arial' })] }),
          new Paragraph({ children: [new TextRun({ text: `Professor: Felipe Vigneron Azevedo  |  Semana: ${semana || ''}`, size: 17, font: 'Arial' })] })
        ], { fill: 'B3E5A0', w: W1 }),
        tc([
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: docName || '', bold: true, size: 18, font: 'Arial' })] })
        ], { fill: 'FFFFFF', w: W2 })
      ]}),
      new TableRow({ children: [
        tc([new Paragraph({ children: [new TextRun({ text: 'Nome do(a) Estudante: _______________________________________________', size: 18, font: 'Arial' })] })], { w: W1 }),
        tc([new Paragraph({ children: [new TextRun({ text: 'Data: ___/___/______', size: 18, font: 'Arial' })] })], { w: W2 })
      ]})
    ]
  });
}
function faixaConfidencial(txt) {
  const label = txt || 'VERSÃO PROFESSOR — CONFIDENCIAL — NÃO DISTRIBUIR AOS ALUNOS';
  return new Table({
    width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
    rows: [new TableRow({ children: [
      tc([new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, bold: true, color: 'FFFFFF', size: 20, font: 'Arial' })] })],
        { fill: '8B0000', w: 9026 })
    ]})]
  });
}
function separador() {
  return [
    p('', { sb: 400, sa: 0, borderTop: true, borderColor: '8B0000' }),
    p('✂  SEPARAR AQUI — A partir daqui: DPM PROFESSOR (CONFIDENCIAL)', { bold: true, color: '8B0000', center: true, sb: 80, sa: 80 }),
    p('', { sb: 0, sa: 400, borderTop: true, borderColor: '8B0000' })
  ];
}
function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing:   { before: opts.sb ?? 120, after: opts.sa ?? 80 },
    border:    opts.borderTop ? { top: { style: BorderStyle.SINGLE, size: 6, color: opts.borderColor || '000000' } } : undefined,
    children:  [new TextRun({ text: text || '', bold: !!opts.bold, italic: !!opts.italic,
      size: opts.size || 20, font: 'Arial', color: opts.color || '000000',
      underline: opts.underline ? {} : undefined })]
  });
}
function parseInlineRuns(text) {
  const runs = [];
  const rx = /\*\*(.+?)\*\*|\*(.+?)\*|([^*]+)/g;
  let m;
  while ((m = rx.exec(text)) !== null) {
    if (m[1])      runs.push(new TextRun({ text: m[1], bold: true,   size: 20, font: 'Arial' }));
    else if (m[2]) runs.push(new TextRun({ text: m[2], italic: true, size: 20, font: 'Arial' }));
    else if (m[3]) runs.push(new TextRun({ text: m[3],               size: 20, font: 'Arial' }));
  }
  return runs.length ? runs : [new TextRun({ text, size: 20, font: 'Arial' })];
}

function makeTableFromMd(rows) {
  // rows: array de arrays de strings (células)
  const BDR_T = { style: BorderStyle.SINGLE, size: 4, color: '2E6B3E' };
  const BDR_L = { style: BorderStyle.SINGLE, size: 2, color: '999999' };
  const BORDERS_HEAD = { top: BDR_T, bottom: BDR_T, left: BDR_L, right: BDR_L };
  const BORDERS_CELL = { top: BDR_L, bottom: BDR_L, left: BDR_L, right: BDR_L };
  const colCount = rows[0].length;
  const colWidth = Math.floor(9026 / colCount);

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: Array(colCount).fill(colWidth),
    rows: rows.map((row, ri) => new TableRow({
      children: row.map(cell => new TableCell({
        borders: ri === 0 ? BORDERS_HEAD : BORDERS_CELL,
        shading: { fill: ri === 0 ? 'E8F5E9' : 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: colWidth, type: WidthType.DXA },
        children: [new Paragraph({
          spacing: { before: 60, after: 60 },
          children: ri === 0
            ? [new TextRun({ text: cell.trim(), bold: true, size: 19, font: 'Arial', color: '1B5E20' })]
            : parseInlineRuns(cell.trim())
        })]
      }))
    }))
  });
}

function mdToDocx(text) {
  if (!text) return [];
  const out   = [];
  const lines = text.split('\n');
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    // Separador --- → ignorar
    if (/^-{3,}$/.test(line.trim())) { i++; continue; }
    // Linha vazia
    if (!line.trim()) { out.push(p('')); i++; continue; }
    // Cabeçalhos
    if (line.startsWith('### ')) { out.push(p(line.slice(4), { bold: true, size: 20, sb: 200, sa: 80  })); i++; continue; }
    if (line.startsWith('## '))  { out.push(p(line.slice(3), { bold: true, size: 22, sb: 240, sa: 100, color: '2E6B3E' })); i++; continue; }
    if (line.startsWith('# '))   { out.push(p(line.slice(2), { bold: true, size: 26, sb: 280, sa: 120 })); i++; continue; }
    // Tabela markdown — detectar bloco
    if (line.startsWith('|')) {
      const tableLines = [];
      while (i < lines.length && lines[i].startsWith('|')) {
        // Ignorar linha separadora |---|---|
        if (!/^\|[-:| ]+\|$/.test(lines[i])) {
          const cells = lines[i].split('|').slice(1, -1);
          tableLines.push(cells);
        }
        i++;
      }
      if (tableLines.length > 0) {
        out.push(p(''));
        out.push(makeTableFromMd(tableLines));
        out.push(p(''));
      }
      continue;
    }
    // Lista
    const indent = line.startsWith('- ');
    const rest   = indent ? line.slice(2) : line;
    out.push(new Paragraph({
      spacing: { before: 80, after: 80 },
      indent:  indent ? { left: 360 } : undefined,
      children: parseInlineRuns(rest)
    }));
    i++;
  }
  return out;
}
function makeDoc(children) {
  return new Document({
    styles:   { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [{ properties: { page: A4_PAGE }, children }]
  });
}

// ── System prompts ─────────────────────────────────────────────────────────
const SYS_NUCLEO = `Assistente pedagógico — IFF Campos Centro, Prof. Felipe Vigneron Azevedo. Método: CLA.
Língua: PT-BR em todas as entregas. Tom: professor experiente, sem marcadores de IA.
Restrições inegociáveis: nunca inventar citações, páginas, datas, títulos ou autores; ancorar cada afirmação no texto; seguir estrutura e ordem pedidas à risca.`;
const SYS_A = SYS_NUCLEO;
const SYS_B = SYS_NUCLEO + '\nModo: engenheiro de prompts — não gera aula. Preservar intenção e tom do autor. Sem inflação de tokens. Apontar explicitamente a melhor alternativa e por quê.';
const SYS_C = SYS_NUCLEO + '\nModo: orientador/banca acadêmico. Devolutiva em 2ª pessoa, sem nota. Rigoroso, propositivo, distância orientador-orientando. Nomear problemas com termo exato. Sem elogio protocolar.';

const GRUPOS_A = `Grupo 1 — Tese: reconstituir o argumento central em mínimo 3 pontos encadeados. Não respondível em uma linha.
Grupo 2 — Mecanismo: identificar COMO o texto constrói o argumento — recursos usados (exemplos, comparações, autoridades, dados) e função de cada um. Respondível apenas com o texto.
Grupo 3 — Tensão: identificar onde o argumento hesita, contradiz premissa própria ou deixa afirmação sem sustentação. A tensão pode ser produtiva — não sugerir que o texto está simplesmente errado.
Grupo 4 — Aplicação: transposição didática a partir dos próprios conceitos do DPM. Modelo: "A partir dos conceitos X e Y, como você, futuro professor, organizaria uma aula de literatura no EM que levasse os alunos a perceber essas categorias em obras que já conhecem? Cite ao menos dois conceitos e descreva brevemente a proposta." Não exigir leitura de obras não transcritas ou anexadas.
Grupo 5 — Implicação: projetar o que se segue necessariamente do argumento — para o campo literário, a prática docente ou textos já estudados. Não requer fontes externas.`;
const GRUPOS_A_FMT_C = 'Grupo 4 (Formato C — substituir): articular o argumento teórico com o texto literário transcrito no DPM Literário desta semana.';
const VERSAO_PROF = `## SP1
## Seção 1 — Questão-Norteadora
Diferente das perguntas de grupo — não aceita respostas iguais ou parecidas com as dos grupos.
Apresentada oralmente antes da leitura do DPM e retomada ao final.
Incluir: texto da questão + sugestão de resposta + páginas do texto.

## Seção 2 — Respostas Esperadas das Questões de Grupos
Tabela: Grupo | Resposta esperada (ancorada no texto com páginas). Uma resposta por grupo.

SUPRIMIR SEMPRE: Equívocos Esperados · Referências Complementares · Perguntas para Aprofundamento · Referência bibliográfica.`;

// ── Prompts ────────────────────────────────────────────────────────────────
function inferirNivel(disc) {
  if (!disc) return 'intermediario';
  const d = disc.toLowerCase();
  if (d.includes('teoria liter') && !d.includes('ii')) return 'iniciante';
  if (d.includes('metodologia')) return 'avancado';
  return 'intermediario';
}
function promptDPMTeorico(inp) {
  const nivelDesc = { iniciante: 'turma iniciante (1º período)', intermediario: 'turma intermediária', avancado: 'turma avançada/pós-graduação' }[inferirNivel(inp.disciplina)];
  const b = inp.budget, fmt = (inp.formato || '').toUpperCase();
  const tempoInicial = fmt === 'B' ? `leitura+discussão ${b.leit || '—'}min` : `conversa norteadora ${b.conv || '—'}min`;
  const grp4ajuste = fmt === 'C' ? '\n' + GRUPOS_A_FMT_C : '';
  return `DPM Teórico — ${inp.disciplina || ''} | Semana ${inp.semana || ''} | Formato ${fmt} | ${inp.nAulas || ''} aulas | ${nivelDesc}
Tema: ${inp.tema || ''} | Textos: ${inp.referencias || ''}${inp.obs ? '\nObs.: ' + inp.obs : ''}
Tempos: ${tempoInicial} · grupo ${b.grp || '—'}min · saída ${b.cpd || '—'}min

Gere APENAS o texto do DPM — sem introduções, sem comentários. Estrutura exata:

== VERSÃO ALUNOS ==
[Referência ABNT completa — antes de qualquer seção numerada. Não repetir ao final.]

## Seção 1 — Corrente Teórica e Contextualização
Com que correntes dialoga, que problema responde, qual método usa. Máx. 3–6 linhas objetivas.

## Seção 2 — Tese Central
2–4 frases. ≥1 citação direta com página (ABNT).

## Seção 3 — Conceitos-Chave
Tabela: Conceito (termo exato + página) | Explicação didática (2–3 frases, linguagem de graduação).
3–7 conceitos em ordem de aparição. Apenas conceitos que o próprio texto define.

## Seção 4 — Parágrafos Centrais do Texto
3–6 citações diretas integrais com ABNT e página.
Cobrir: conceitos centrais · argumento principal · tensões · diálogo com outros autores · conclusão.
NÃO parafrasear.

## Seção 5 — Perguntas de Grupo
${GRUPOS_A}${grp4ajuste}

${VERSAO_PROF}`;
}
function promptDPMLiterario(inp) {
  const nivelDesc = { iniciante: 'turma iniciante', intermediario: 'turma intermediária', avancado: 'turma avançada' }[inferirNivel(inp.disciplina)];
  const b = inp.budget, fmt = (inp.formato || '').toUpperCase();
  const isMeto = (inp.disciplina || '').toLowerCase().includes('metodologia');
  const tipo   = isMeto ? 'texto demonstrativo (não literário)' : 'texto literário';
  const grp4   = fmt === 'C' ? 'Grupo 4 — Aplicação: articular com o DPM Teórico desta semana.' : 'Grupo 4 — Intertexto: relações com outros textos evidenciadas pelo próprio texto.';
  return `DPM ${isMeto ? 'Demonstrativo' : 'Literário'} — ${inp.disciplina || ''} | Semana ${inp.semana || ''} | Formato ${fmt} | ${nivelDesc}
Tema: ${inp.tema || ''} | Textos: ${inp.referencias || ''}${inp.obs ? '\nObs.: ' + inp.obs : ''}
Tempos: discussão ${b.lit || b.leit || '—'}min · grupo ${b.grp || '—'}min · saída ${b.cpd || '—'}min

Gere APENAS o texto do DPM — sem introduções, sem comentários. Estrutura exata:

== VERSÃO ALUNOS ==
[Referência ABNT completa — antes de qualquer seção numerada.]

## Seção 1 — Tese Central do ${tipo}
2–4 frases com o argumento central.

## Seção 2 — Forma
Gênero · estrutura · narrador/voz · tempo · espaço · dicção${isMeto ? ' · tipo de argumento · metodologia demonstrada' : ''}.

## Seção 3 — Conteúdo
Temas · personagens/agentes · conflito central · desfecho.

## Seção 4 — Contexto
Contexto histórico-literário${isMeto ? '/acadêmico' : ''} · autor · período.

## Seção 5 — Intertexto
Relações com outros textos evidenciadas pelo próprio texto.

## Seção 6 — Parágrafos Centrais do Texto
3–5 citações diretas integrais com ABNT e página.

## Seção 7 — Perguntas de Grupo
Grupo 1 — Forma · Grupo 2 — Conteúdo · Grupo 3 — Contexto histórico-literário
${grp4}
Grupo 5 — Lacuna: indicar o que o DPM não cobre e orientar a consultar o texto original.

${VERSAO_PROF}`;
}
function promptQuiz(inp) {
  const b = inp.budget;
  return `Quiz de 10 min — ${inp.disciplina || ''} | Semana ${inp.semana || ''} | Tema: ${inp.tema || ''}
Textos: ${inp.referencias || ''} | Tempo na aula: ${b.quiz || 10}min

Gere APENAS as 5 questões e o gabarito — sem cabeçalho, sem introdução, sem faixa.

5 questões de múltipla escolha (A–D).
Formato: número + enunciado em negrito → alternativas A/B/C/D → linha em branco.
Q1: reformulação da questão-norteadora do DPM (mesma temática, formulação diferente).
Q2–Q5: baseadas no texto e no DPM, sem coincidir com perguntas dos grupos.

Linha divisória, depois:
GABARITO — VERSÃO PROFESSOR — NÃO DISTRIBUIR AOS ALUNOS
Uma linha por questão: número · resposta · comentário breve · (SOBRENOME, ano, p. X).`;
}
function promptBimestral(inp) {
  return `Questões Bimestrais — ${inp.disciplina || ''} | Semana ${inp.semana || ''} | Textos: ${inp.referencias || ''}

Gere APENAS as 2 questões e o gabarito — sem cabeçalho, sem faixa, sem introdução.

Q1: indicação "(nível: compreensão)" → enunciado em negrito → alternativas A/B/C/D → linha em branco.
Q2: indicação "(nível: interpretação)" → enunciado em negrito → alternativas A/B/C/D → linha em branco.

Linha divisória, depois:
GABARITO — VERSÃO PROFESSOR
Uma linha por questão: número · resposta · comentário breve · (SOBRENOME, ano, p. X).`;
}

// ── Geradores ──────────────────────────────────────────────────────────────
function toFileContent(files) {
  return (files || []).map(f => ({
    type: 'document',
    source: { type: 'base64', media_type: f.media_type, data: f.data }
  }));
}
async function gerarDPM(inp, files, tipo) {
  const prompt = tipo === 'teorico' ? promptDPMTeorico(inp) : promptDPMLiterario(inp);
  const res    = await client.messages.create({
    model: MODEL, max_tokens: 4096, system: SYS_A,
    messages: [{ role: 'user', content: [...toFileContent(files), { type: 'text', text: prompt }] }]
  });
  const text      = res.content[0].text;
  const splitIdx  = text.indexOf('## SP1');
  const alunosRaw = splitIdx > 0 ? text.slice(0, splitIdx) : text;
  const profRaw   = splitIdx > 0 ? text.slice(splitIdx)    : '';
  const alunosText = alunosRaw.replace(/^==\s*VERS[ÃA]O ALUNOS\s*==\s*/im, '').trim();
  const profText   = profRaw.replace(/^##\s*SP1\s*/m, '').trim();
  const isMeto  = (inp.disciplina || '').toLowerCase().includes('metodologia');
  const docName = tipo === 'teorico' ? 'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Teórico'
    : isMeto ? 'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Demonstrativo'
             : 'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Literário';
  const label = tipo === 'teorico' ? 'DPM TEÓRICO' : isMeto ? 'DPM DEMONSTRATIVO' : 'DPM LITERÁRIO';
  const children = [
    makeHeader(inp.disciplina, inp.semana, docName),
    p(''), p(`${label} — VERSÃO ALUNOS`, { bold: true, size: 24, sb: 200 }), p(''),
    ...mdToDocx(alunosText)
  ];
  if (profText) children.push(...separador(), faixaConfidencial(), p(''), ...mdToDocx(profText));
  return Packer.toBase64String(makeDoc(children));
}
async function gerarQuizDoc(inp, files) {
  const res = await client.messages.create({
    model: MODEL, max_tokens: 2048, system: SYS_A,
    messages: [{ role: 'user', content: [...toFileContent(files), { type: 'text', text: promptQuiz(inp) }] }]
  });
  return Packer.toBase64String(makeDoc([makeHeader(inp.disciplina, inp.semana, 'QUIZ DE 10 MINUTOS'), p(''), ...mdToDocx(res.content[0].text)]));
}
async function gerarBimestralDoc(inp, files) {
  const res = await client.messages.create({
    model: MODEL, max_tokens: 1024, system: SYS_A,
    messages: [{ role: 'user', content: [...toFileContent(files), { type: 'text', text: promptBimestral(inp) }] }]
  });
  return Packer.toBase64String(makeDoc([faixaConfidencial('QUESTÕES BIMESTRAIS — VERSÃO PROFESSOR — NÃO DISTRIBUIR AOS ALUNOS'), p(''), ...mdToDocx(res.content[0].text)]));
}
async function gerarOtimizador(inp) {
  const res = await client.messages.create({
    model: MODEL, max_tokens: 2048, system: SYS_B,
    messages: [{ role: 'user', content: `${inp.origem === 'painel' ? 'Material do Painel CLA.\n' : ''}Queixa/objetivo: ${inp.queixa || 'não especificado'}\n\n${inp.prompt}\n\nEntregar: 1) DIAGNÓSTICO 2) VERSÃO OTIMIZADA 3) O QUE MUDOU E POR QUÊ` }]
  });
  return res.content[0].text;
}
async function gerarDevolutiva(inp, files) {
  const papel = inp.papel === 'banca' ? 'banca (avaliativa)' : 'orientador (formativa)';
  const fase  = { inicio: 'início', andamento: 'andamento', concluido: 'concluído (pré-banca)' }[inp.fase] || '';
  const nivel = { artigo: 'artigo/TCC', dissertacao: 'dissertação', tese: 'tese' }[inp.nivel] || inp.nivel;
  const res   = await client.messages.create({
    model: MODEL, max_tokens: 3000, system: SYS_C,
    messages: [{ role: 'user', content: [...toFileContent(files), { type: 'text', text:
`Papel: ${papel}${inp.papel !== 'banca' ? ' · Fase: ' + fase : ''} · Nível: ${nivel}
${inp.foco ? 'Foco: ' + inp.foco : ''}
${inp.contexto ? 'Contexto/trechos:\n' + inp.contexto : ''}
${files && files.length ? 'Trabalho anexado acima.' : 'Usar contexto/trechos fornecidos.'}

Critérios (por peso): 1) Cumprimento dos objetivos 2) Originalidade 3) Fundamentação teórica 4) Correção conceitual · Clareza · Consistência · ABNT

Estrutura:
1 — Leitura geral (2–4 frases)
2 — Pontos por critério (só onde há algo a dizer)
3 — Apontamentos cirúrgicos: trecho → problema → sugestão → fonte
4 — Prioridades (2–3 providências)
5 — Próximo passo` }] }]
  });
  return res.content[0].text;
}
async function gerarDevolutivaDoc(inp, files) {
  const text  = await gerarDevolutiva(inp, files);
  const papel = inp.papel === 'banca' ? 'Banca' : 'Orientador';
  const fase  = { inicio: 'Início', andamento: 'Andamento', concluido: 'Concluído' }[inp.fase] || '';
  const nivel = { artigo: 'Artigo/TCC', dissertacao: 'Dissertação', tese: 'Tese' }[inp.nivel] || inp.nivel;
  return Packer.toBase64String(makeDoc([
    p(`Devolutiva · ${papel}${fase ? ' · ' + fase : ''} · ${nivel}`, { size: 18, color: '555555', sb: 0 }),
    p(''), ...mdToDocx(text)
  ]));
}

// ── Middleware ─────────────────────────────────────────────────────────────
app.use(cookieParser());
app.use(express.json({ limit: '50mb' }));

// CORS
app.use((req, res, next) => {
  const origin = req.headers.origin;
  const allowed = process.env.SITE_URL || origin || '*';
  res.setHeader('Access-Control-Allow-Origin', allowed);
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// ── Rotas de autenticação ──────────────────────────────────────────────────
app.get('/auth-check', (req, res) => {
  const token = req.cookies?.cla_session;
  const data  = verifyToken(token);
  if (data) return res.json({ ok: true, email: data.email, authUrl: null });
  const params = new URLSearchParams({
    client_id:     requireEnv('GOOGLE_CLIENT_ID'),
    redirect_uri:  requireEnv('REDIRECT_URI'),
    response_type: 'code',
    scope:         'openid email profile',
    prompt:        'select_account'
  });
  res.json({ ok: false, email: null, authUrl: `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}` });
});

app.get('/auth-callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send('Código ausente.');
  try {
    const tokens = await httpsPost('https://oauth2.googleapis.com/token', {
      code,
      client_id:     requireEnv('GOOGLE_CLIENT_ID'),
      client_secret: requireEnv('GOOGLE_CLIENT_SECRET'),
      redirect_uri:  requireEnv('REDIRECT_URI'),
      grant_type:    'authorization_code'
    });
    if (!tokens.access_token) return res.status(401).send('Falha na autenticação Google.');
    const userInfo = await httpsGet('https://www.googleapis.com/oauth2/v2/userinfo', tokens.access_token);
    const allowed  = requireEnv('ALLOWED_EMAIL');
    if (userInfo.email !== allowed) {
      return res.status(403).send(`<html><body style="font-family:sans-serif;text-align:center;padding:60px">
        <h2>Acesso negado</h2>
        <p>Este painel é de uso exclusivo do Prof. Felipe Vigneron Azevedo.</p>
        <p style="color:#888">${userInfo.email}</p>
      </body></html>`);
    }
    res.cookie('cla_session', makeToken(userInfo.email), {
      httpOnly: true, secure: true, sameSite: 'lax', maxAge: 86400000 * 7, path: '/'
    });
    res.redirect('/');
  } catch (err) {
    console.error('auth-callback:', err);
    res.status(500).send('Erro interno.');
  }
});

// ── Rota principal da API ──────────────────────────────────────────────────
app.post('/cla-api', requireSession, upload.array('files'), async (req, res) => {
  try {
    // Parsear payload JSON
    const pl     = JSON.parse(req.body.payload || '{}');
    const mode   = pl.mode;
    const task   = pl.task;
    const inputs = pl.inputs || {};
    // Arquivos via multer (memória)
    const files  = (req.files || []).map(f => ({
      name:       f.originalname,
      media_type: f.mimetype || 'application/pdf',
      data:       f.buffer.toString('base64')
    }));

    const inp = inputs;
    const fmt = (inp.formato || '').toUpperCase();
    inp.budget = (BUDGET[fmt] || {})[inp.nAulas || ''] || {};

    console.log(`cla-api: mode=${mode} task=${task} arquivos=${files.length}`);

    if (mode === 'A') {
      const sem = inp.semana || 'X';
      if (task === 'dpm_teorico')   return res.json({ docx: await gerarDPM(inp, files, 'teorico'),   filename: `DPM_Teorico_Sem${sem}.docx`,   warnings: [] });
      if (task === 'dpm_literario') return res.json({ docx: await gerarDPM(inp, files, 'literario'), filename: `DPM_Literario_Sem${sem}.docx`, warnings: [] });
      if (task === 'quiz')          return res.json({ docx: await gerarQuizDoc(inp, files),           filename: `Quiz_Sem${sem}.docx`,          warnings: [] });
      if (task === 'bimestral')     return res.json({ docx: await gerarBimestralDoc(inp, files),      filename: `Bimestral_Sem${sem}.docx`,     warnings: [] });
      return res.status(400).json({ error: `Tarefa desconhecida: ${task}` });
    }
    if (mode === 'B') {
      if (!inp.prompt) return res.status(400).json({ error: 'Campo prompt obrigatório.' });
      return res.json({ text: await gerarOtimizador(inp), model: MODEL, warnings: [] });
    }
    if (mode === 'C') {
      if (task === 'chat') return res.json({ text: await gerarDevolutiva(inp, files || []), model: MODEL, warnings: [] });
      return res.json({ docx: await gerarDevolutivaDoc(inp, files || []), filename: `Devolutiva_${inp.nivel || 'trabalho'}.docx`, warnings: [] });
    }
    return res.status(400).json({ error: `Modo desconhecido: ${mode}` });
  } catch (e) {
    console.error('cla-api:', e);
    res.status(500).json({ error: e.message || 'Erro interno.' });
  }
});

// ── Servir frontend ────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  const indexPath = path.join(__dirname, 'index.html');
  console.log('Serving index.html from:', indexPath);
  res.sendFile(indexPath);
});

// Fallback para qualquer rota não encontrada
app.use((req, res) => {
  console.log('404:', req.url);
  res.status(404).send('Not found: ' + req.url);
});

app.listen(PORT, () => console.log(`Painel CLA rodando na porta ${PORT}`));
