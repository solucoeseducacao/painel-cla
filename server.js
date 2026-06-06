'use strict';
const express=require('express'),cookieParser=require('cookie-parser'),multer=require('multer'),https=require('https'),crypto=require('crypto'),path=require('path'),Anthropic=require('@anthropic-ai/sdk'),{Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,AlignmentType,BorderStyle,WidthType,ShadingType}=require('docx');
const app=express(),upload=multer({storage:multer.memoryStorage()}),PORT=process.env.PORT||3000;
function env(n){const v=process.env[n];if(!v)throw new Error(`Env missing: ${n}`);return v}
function sign(d){return crypto.createHmac('sha256',env('SESSION_SECRET')).update(d).digest('base64url')}
function makeToken(e){const p=Buffer.from(JSON.stringify({email:e,exp:Date.now()+86400000*7})).toString('base64url');return`${p}.${sign(p)}`}
function verifyToken(t){if(!t)return null;const i=t.indexOf('.');if(i<0)return null;const p=t.slice(0,i),s=t.slice(i+1);if(!p||!s)return null;try{const e=sign(p),eb=Buffer.from(e,'base64url'),sb=Buffer.from(s,'base64url');if(eb.length!==sb.length||!crypto.timingSafeEqual(eb,sb))return null;const d=JSON.parse(Buffer.from(p,'base64url').toString());if(d.exp<Date.now()||d.email!==env('ALLOWED_EMAIL'))return null;return d}catch{return null}}
function requireSession(req,res,next){const d=verifyToken(req.cookies?.cla_session);if(!d)return res.status(401).json({error:'Sessão inválida ou expirada.'});req.user=d;next()}
function httpsPost(url,data){return new Promise((resolve,reject)=>{const b=new URLSearchParams(data).toString(),u=new URL(url),r=https.request({hostname:u.hostname,path:u.pathname,method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(b)}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>{try{resolve(JSON.parse(d))}catch(e){reject(e)}})});r.on('error',reject);r.write(b);r.end()})}
function httpsGet(url,token){return new Promise((resolve,reject)=>{const u=new URL(url),r=https.request({hostname:u.hostname,path:u.pathname+u.search,method:'GET',headers:{Authorization:`Bearer ${token}`}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>{try{resolve(JSON.parse(d))}catch(e){reject(e)}})});r.on('error',reject);r.end()})}
const client=new Anthropic({apiKey:process.env.ANTHROPIC_API_KEY}),MODEL='claude-sonnet-4-6',MODEL_FAST='claude-haiku-4-5';
const BUDGET={A:{'2':{conv:20,quiz:10,entrada:10,grp:20,reg:15,cpd:15},'3':{conv:30,quiz:10,entrada:10,grp:35,reg:20,cpd:15},'4':{conv:45,quiz:10,entrada:15,grp:50,reg:25,cpd:15}},B:{'2':{leit:30,quiz:10,grp:20,reg:15,cpd:15},'3':{leit:45,quiz:10,grp:35,reg:15,cpd:15},'4':{leit:65,quiz:10,grp:50,reg:20,cpd:15}},C:{'2':{conv:15,lit:15,quiz:10,grp:20,reg:15,cpd:15},'3':{conv:20,lit:20,quiz:10,grp:35,reg:20,cpd:15},'4':{conv:30,lit:30,quiz:10,grp:45,reg:25,cpd:20}}};
const A4={size:{width:11906,height:16838},margin:{top:1440,right:1440,bottom:1440,left:1440}};
const BDR={style:BorderStyle.SINGLE,size:1,color:'999999'},BORDERS={top:BDR,bottom:BDR,left:BDR,right:BDR};
function tc(children,opts={}){return new TableCell({borders:BORDERS,shading:{fill:opts.fill||'FFFFFF',type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},width:{size:opts.w||4513,type:WidthType.DXA},children})}
function makeHeader(disc,semana,doc){const W1=5416,W2=3610;return new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[W1,W2],rows:[new TableRow({children:[tc([new Paragraph({children:[new TextRun({text:'INSTITUTO FEDERAL FLUMINENSE',bold:true,size:22,font:'Arial'})]}),new Paragraph({children:[new TextRun({text:'Campus Campos Centro · Licenciatura em Letras — Português e Literaturas',size:21,font:'Arial'})]}),new Paragraph({children:[new TextRun({text:`Disciplina: ${disc||''}`,size:21,font:'Arial'})]}),new Paragraph({children:[new TextRun({text:`Professor: Felipe Vigneron Azevedo  |  Semana: ${semana||''}`,size:21,font:'Arial'})]})],{fill:'B3E5A0',w:W1}),tc([new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:doc||'',bold:true,size:22,font:'Arial'})]})],{fill:'FFFFFF',w:W2})]}),new TableRow({children:[tc([new Paragraph({children:[new TextRun({text:'Nome do(a) Estudante: _______________________________________________',size:22,font:'Arial'})]})],{w:W1}),tc([new Paragraph({children:[new TextRun({text:'Data: ___/___/______',size:22,font:'Arial'})]})],{w:W2})]})]})}
function faixaConf(txt){const l=txt||'VERSÃO PROFESSOR — CONFIDENCIAL — NÃO DISTRIBUIR AOS ALUNOS';return new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],rows:[new TableRow({children:[tc([new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:l,bold:true,color:'FFFFFF',size:24,font:'Arial'})]})],{fill:'8B0000',w:9026})]})]})}
function separador(){return[p('',{sb:400,sa:0,borderTop:true,borderColor:'8B0000'}),p('✂  SEPARAR AQUI — A partir daqui: DPM PROFESSOR (CONFIDENCIAL)',{bold:true,color:'8B0000',center:true,sb:80,sa:80}),p('',{sb:0,sa:400,borderTop:true,borderColor:'8B0000'})]}
function p(text,opts={}){return new Paragraph({alignment:opts.center?AlignmentType.CENTER:AlignmentType.LEFT,spacing:{before:opts.sb??120,after:opts.sa??80},border:opts.borderTop?{top:{style:BorderStyle.SINGLE,size:6,color:opts.borderColor||'000000'}}:undefined,children:[new TextRun({text:text||'',bold:!!opts.bold,italic:!!opts.italic,size:opts.size||24,font:'Arial',color:opts.color||'000000',underline:opts.underline?{}:undefined})]})}
function parseInlineRuns(text){const r=[],rx=/\*\*(.+?)\*\*|\*(.+?)\*|([^*]+)/g;let m;while((m=rx.exec(text))!==null){if(m[1])r.push(new TextRun({text:m[1],bold:true,size:24,font:'Arial'}));else if(m[2])r.push(new TextRun({text:m[2],italic:true,size:24,font:'Arial'}));else if(m[3])r.push(new TextRun({text:m[3],size:24,font:'Arial'}))}return r.length?r:[new TextRun({text,size:24,font:'Arial'})]}
function makeTableFromMd(rows){const colC=rows[0].length,colW=Math.floor(9026/colC);return new Table({width:{size:9026,type:WidthType.DXA},columnWidths:Array(colC).fill(colW),rows:rows.map((row,ri)=>new TableRow({children:row.map(cell=>new TableCell({borders:ri===0?{top:{style:BorderStyle.SINGLE,size:4,color:'2E6B3E'},bottom:{style:BorderStyle.SINGLE,size:4,color:'2E6B3E'},left:{style:BorderStyle.SINGLE,size:2,color:'999999'},right:{style:BorderStyle.SINGLE,size:2,color:'999999'}}:{top:{style:BorderStyle.SINGLE,size:2,color:'999999'},bottom:{style:BorderStyle.SINGLE,size:2,color:'999999'},left:{style:BorderStyle.SINGLE,size:2,color:'999999'},right:{style:BorderStyle.SINGLE,size:2,color:'999999'}},shading:{fill:ri===0?'E8F5E9':'FFFFFF',type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},width:{size:colW,type:WidthType.DXA},children:[new Paragraph({spacing:{before:60,after:60},children:ri===0?[new TextRun({text:cell.trim(),bold:true,size:23,font:'Arial',color:'1B5E20'})]:parseInlineRuns(cell.trim())})]}))}))})}
function mdToDocx(text){if(!text)return[];const out=[],lines=text.split('\n');let i=0;while(i<lines.length){const line=lines[i];if(/^-{3,}$/.test(line.trim())){i++;continue}if(!line.trim()){out.push(p(''));i++;continue}if(line.startsWith('### ')){out.push(p(line.slice(4),{bold:true,size:28,sb:200,sa:80}));i++;continue}if(line.startsWith('## ')){out.push(p(line.slice(3),{bold:true,size:26,sb:240,sa:100,color:'2E6B3E'}));i++;continue}if(line.startsWith('# ')){out.push(p(line.slice(2),{bold:true,size:30,sb:280,sa:120}));i++;continue}if(line.startsWith('|')){const tbl=[];while(i<lines.length&&lines[i].startsWith('|')){if(!/^\|[-:| ]+\|$/.test(lines[i])){const cells=lines[i].split('|').slice(1,-1);tbl.push(cells)}i++}if(tbl.length>0){out.push(p(''));out.push(makeTableFromMd(tbl));out.push(p(''))}continue}const indent=line.startsWith('- '),rest=indent?line.slice(2):line;out.push(new Paragraph({spacing:{before:80,after:80},indent:indent?{left:360}:undefined,children:parseInlineRuns(rest)}));i++}return out}
function makeDoc(children){return new Document({styles:{default:{document:{run:{font:'Arial',size:24}}}},sections:[{properties:{page:A4},children}]})}
const SYS_A=`Assistente pedagógico — IFF Campos Centro, Prof. Felipe Vigneron Azevedo. Método: CLA.
Língua: PT-BR em todas as entregas. Tom: professor experiente, sem marcadores de IA.
Restrições inegociáveis: nunca inventar citações, páginas, datas, títulos ou autores; ancorar cada afirmação no texto; seguir estrutura e ordem pedidas à risca.`;
const SYS_B=SYS_A+'\nModo: engenheiro de prompts — não gera aula. Preservar intenção e tom do autor. Sem inflação de tokens. Apontar explicitamente a melhor alternativa e por quê.';
const SYS_C=SYS_A+'\nModo: orientador/banca acadêmico. Devolutiva em 2ª pessoa, sem nota. Rigoroso, propositivo, distância orientador-orientando. Nomear problemas com termo exato. Sem elogio protocolar.';
const GRUPOS_A=`Grupo 1 — Tese: reconstituir o argumento central em mínimo 3 pontos encadeados. Não respondível em uma linha.
Grupo 2 — Mecanismo: identificar COMO o texto constrói o argumento — recursos usados e função de cada um.
Grupo 3 — Tensão: identificar onde o argumento hesita, contradiz premissa ou deixa afirmação sem sustentação.
Grupo 4 — Aplicação: transposição didática a partir dos conceitos do DPM. Modelo: "A partir dos conceitos X e Y, como você organizaria uma aula de literatura no EM que levasse os alunos a perceber essas categorias em obras conhecidas? Cite ao menos dois conceitos."
Grupo 5 — Implicação: projetar o que se segue do argumento para o campo literário ou prática docente.`;
const GRUPOS_A_FMT_C='Grupo 4 (Formato C — substituir): articular o argumento teórico com o texto literário transcrito no DPM Literário desta semana.';
const VERSAO_PROF=`## SP1
## Seção 1 — Questão-Norteadora
Diferente das perguntas de grupo — não aceita respostas iguais. Apresentada oralmente antes da leitura do DPM e retomada ao final. Incluir: texto da questão + sugestão de resposta + páginas.

## Seção 2 — Respostas Esperadas das Questões de Grupos
Tabela: Grupo | Resposta esperada (ancorada no texto com páginas). Uma resposta por grupo.
SUPRIMIR SEMPRE: Equívocos Esperados · Referências Complementares · Perguntas para Aprofundamento · Referência bibliográfica.`;
function inferirNivel(disc){if(!disc)return'intermediario';const d=disc.toLowerCase();if(d.includes('teoria liter')&&!d.includes('ii'))return'iniciante';if(d.includes('metodologia'))return'avancado';return'intermediario'}
function promptDPMTeorico(inp){const nd={iniciante:'turma iniciante (1º período)',intermediario:'turma intermediária',avancado:'turma avançada/pós-graduação'}[inferirNivel(inp.disciplina)];const b=inp.budget,f=(inp.formato||'').toUpperCase(),ti=f==='B'?`leitura+discussão ${b.leit||'—'}min`:`conversa norteadora ${b.conv||'—'}min`,g4=f==='C'?'\n'+GRUPOS_A_FMT_C:'';return`DPM Teórico — ${inp.disciplina||''} | Semana ${inp.semana||''} | Formato ${f} | ${inp.nAulas||''} aulas | ${nd}
Tema: ${inp.tema||''} | Textos: ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}
Tempos: ${ti} · grupo ${b.grp||'—'}min · saída ${b.cpd||'—'}min

Gere APENAS o texto do DPM — sem introduções, sem comentários. Estrutura exata:

== VERSÃO ALUNOS ==
[Referência ABNT completa — antes de qualquer seção numerada.]

## Seção 1 — Corrente Teórica e Contextualização
Com que correntes dialoga, que problema responde, qual método usa. Máx. 3–6 linhas.

## Seção 2 — Tese Central
2–4 frases. ≥1 citação direta com página (ABNT).

## Seção 3 — Conceitos-Chave
Tabela: Conceito (termo exato + página) | Explicação didática (2–3 frases).
3–7 conceitos em ordem de aparição. Apenas conceitos que o próprio texto define.

## Seção 4 — Parágrafos Centrais do Texto
3–6 citações diretas integrais com ABNT e página.
Cobrir: conceitos centrais · argumento principal · tensões · diálogo com outros autores · conclusão.
NÃO parafrasear.

## Seção 5 — Perguntas de Grupo
${GRUPOS_A}${g4}

${VERSAO_PROF}`}
function promptDPMLiterario(inp){const nd={iniciante:'turma iniciante',intermediario:'turma intermediária',avancado:'turma avançada'}[inferirNivel(inp.disciplina)];const b=inp.budget,f=(inp.formato||'').toUpperCase(),isM=(inp.disciplina||'').toLowerCase().includes('metodologia'),tipo=isM?'texto demonstrativo (não literário)':'texto literário',g4=f==='C'?'Grupo 4 — Aplicação: articular com o DPM Teórico desta semana.':'Grupo 4 — Intertexto: relações com outros textos evidenciadas pelo próprio texto.';return`DPM ${isM?'Demonstrativo':'Literário'} — ${inp.disciplina||''} | Semana ${inp.semana||''} | Formato ${f} | ${nd}
Tema: ${inp.tema||''} | Textos: ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}
Tempos: discussão ${b.lit||b.leit||'—'}min · grupo ${b.grp||'—'}min · saída ${b.cpd||'—'}min

Gere APENAS o texto do DPM — sem introduções. Estrutura exata:

== VERSÃO ALUNOS ==
[Referência ABNT completa — antes de qualquer seção numerada.]

## Seção 1 — Tese Central do ${tipo}
2–4 frases com o argumento central.

## Seção 2 — Forma
Gênero · estrutura · narrador/voz · tempo · espaço · dicção${isM?' · tipo de argumento · metodologia demonstrada':''}.

## Seção 3 — Conteúdo
Temas · personagens/agentes · conflito central · desfecho.

## Seção 4 — Contexto
Contexto histórico-literário${isM?'/acadêmico':''} · autor · período.

## Seção 5 — Intertexto
Relações com outros textos evidenciadas pelo próprio texto.

## Seção 6 — Parágrafos Centrais do Texto
3–5 citações diretas integrais com ABNT e página.

## Seção 7 — Perguntas de Grupo
Grupo 1 — Forma · Grupo 2 — Conteúdo · Grupo 3 — Contexto histórico-literário
${g4}
Grupo 5 — Lacuna: indicar o que o DPM não cobre e orientar a consultar o texto original.

${VERSAO_PROF}`}
function promptQuiz(inp){const b=inp.budget;return`Quiz de 10 min — ${inp.disciplina||''} | Semana ${inp.semana||''} | Tema: ${inp.tema||''}
Textos: ${inp.referencias||''} | Tempo na aula: ${b.quiz||10}min

Gere APENAS as 5 questões e o gabarito — sem cabeçalho, sem faixa.

5 questões de múltipla escolha (A–D).
Formato: número + enunciado em negrito → alternativas A/B/C/D → linha em branco.
Q1: reformulação da questão-norteadora do DPM (mesma temática, formulação diferente).
Q2–Q5: baseadas no texto e no DPM, sem coincidir com perguntas dos grupos.

Linha divisória, depois:
GABARITO — VERSÃO PROFESSOR — NÃO DISTRIBUIR AOS ALUNOS
Uma linha por questão: número · resposta · comentário breve · (SOBRENOME, ano, p. X).`}
function promptBimestral(inp){return`Questões Bimestrais — ${inp.disciplina||''} | Semana ${inp.semana||''} | Textos: ${inp.referencias||''}

Gere APENAS as 2 questões e o gabarito — sem cabeçalho.

Q1: indicação "(nível: compreensão)" → enunciado em negrito → alternativas A/B/C/D → linha em branco.
Q2: indicação "(nível: interpretação)" → enunciado em negrito → alternativas A/B/C/D → linha em branco.

Linha divisória, depois:
GABARITO — VERSÃO PROFESSOR
Uma linha por questão: número · resposta · comentário breve · (SOBRENOME, ano, p. X).`}
function toFileContent(files){return(files||[]).map(f=>({type:'document',source:{type:'base64',media_type:f.media_type,data:f.data}}))}
async function gerarDPM(inp,files,tipo){const prompt=tipo==='teorico'?promptDPMTeorico(inp):promptDPMLiterario(inp);const res=await client.messages.create({model:MODEL,max_tokens:3000,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:prompt}]}]});const text=res.content[0].text,si=text.indexOf('## SP1'),aR=si>0?text.slice(0,si):text,pR=si>0?text.slice(si):'',aT=aR.replace(/^==\s*VERS[ÃA]O ALUNOS\s*==\s*/im,'').trim(),pT=pR.replace(/^##\s*SP1\s*/m,'').trim(),isM=(inp.disciplina||'').toLowerCase().includes('metodologia'),docName=tipo==='teorico'?'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Teórico':isM?'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Demonstrativo':'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Literário',label=tipo==='teorico'?'DPM TEÓRICO':isM?'DPM DEMONSTRATIVO':'DPM LITERÁRIO';const children=[makeHeader(inp.disciplina,inp.semana,docName),p(''),p(`${label} — VERSÃO ALUNOS`,{bold:true,size:28,sb:200}),p(''),...mdToDocx(aT)];if(pT)children.push(...separador(),faixaConf(),p(''),...mdToDocx(pT));return Packer.toBase64String(makeDoc(children))}
async function gerarQuizDoc(inp,files){const res=await client.messages.create({model:MODEL_FAST,max_tokens:1200,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptQuiz(inp)}]}]});return Packer.toBase64String(makeDoc([makeHeader(inp.disciplina,inp.semana,'QUIZ DE 10 MINUTOS'),p(''),...mdToDocx(res.content[0].text)]))}
async function gerarBimestralDoc(inp,files){const res=await client.messages.create({model:MODEL_FAST,max_tokens:800,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptBimestral(inp)}]}]});return Packer.toBase64String(makeDoc([faixaConf('QUESTÕES BIMESTRAIS — VERSÃO PROFESSOR — NÃO DISTRIBUIR AOS ALUNOS'),p(''),...mdToDocx(res.content[0].text)]))}
async function gerarOtimizador(inp){const res=await client.messages.create({model:MODEL,max_tokens:1500,system:SYS_B,messages:[{role:'user',content:`${inp.origem==='painel'?'Material do Painel CLA.\n':''}Queixa/objetivo: ${inp.queixa||'não especificado'}\n\n${inp.prompt}\n\nEntregar: 1) DIAGNÓSTICO 2) VERSÃO OTIMIZADA 3) O QUE MUDOU E POR QUÊ`}]});return res.content[0].text}
async function gerarDevolutiva(inp,files){const papel=inp.papel==='banca'?'banca (avaliativa)':'orientador (formativa)',fase={inicio:'início',andamento:'andamento',concluido:'concluído (pré-banca)'}[inp.fase]||'',nivel={artigo:'artigo/TCC',dissertacao:'dissertação',tese:'tese'}[inp.nivel]||inp.nivel;const res=await client.messages.create({model:MODEL,max_tokens:2500,system:SYS_C,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:`Papel: ${papel}${inp.papel!=='banca'?' · Fase: '+fase:''} · Nível: ${nivel}\n${inp.foco?'Foco: '+inp.foco:''}\n${inp.contexto?'Contexto/trechos:\n'+inp.contexto:''}\n${files&&files.length?'Trabalho anexado acima.':'Usar contexto/trechos fornecidos.'}\n\nCritérios (por peso): 1) Cumprimento dos objetivos 2) Originalidade 3) Fundamentação teórica 4) Correção conceitual · Clareza · Consistência · ABNT\n\nEstrutura:\n1 — Leitura geral (2–4 frases)\n2 — Pontos por critério (só onde há algo a dizer)\n3 — Apontamentos cirúrgicos: trecho → problema → sugestão → fonte\n4 — Prioridades (2–3 providências)\n5 — Próximo passo`}]}]});return res.content[0].text}
async function gerarDevolutivaDoc(inp,files){const text=await gerarDevolutiva(inp,files),papel=inp.papel==='banca'?'Banca':'Orientador',fase={inicio:'Início',andamento:'Andamento',concluido:'Concluído'}[inp.fase]||'',nivel={artigo:'Artigo/TCC',dissertacao:'Dissertação',tese:'Tese'}[inp.nivel]||inp.nivel;return Packer.toBase64String(makeDoc([p(`Devolutiva · ${papel}${fase?' · '+fase:''} · ${nivel}`,{size:18,color:'555555',sb:0}),p(''),...mdToDocx(text)]))}
app.use(cookieParser());app.use(express.json({limit:'50mb'}));
app.use((req,res,next)=>{const o=req.headers.origin,a=process.env.SITE_URL||o||'*';res.setHeader('Access-Control-Allow-Origin',a);res.setHeader('Access-Control-Allow-Credentials','true');res.setHeader('Access-Control-Allow-Methods','GET, POST, OPTIONS');res.setHeader('Access-Control-Allow-Headers','Content-Type');if(req.method==='OPTIONS')return res.sendStatus(200);next()});
app.get('/auth-check',(req,res)=>{const d=verifyToken(req.cookies?.cla_session);if(d)return res.json({ok:true,email:d.email,authUrl:null});const p=new URLSearchParams({client_id:env('GOOGLE_CLIENT_ID'),redirect_uri:env('REDIRECT_URI'),response_type:'code',scope:'openid email profile',prompt:'select_account'});res.json({ok:false,email:null,authUrl:`https://accounts.google.com/o/oauth2/v2/auth?${p.toString()}`})});
app.get('/auth-callback',async(req,res)=>{const c=req.query.code;if(!c)return res.status(400).send('Código ausente.');try{const t=await httpsPost('https://oauth2.googleapis.com/token',{code:c,client_id:env('GOOGLE_CLIENT_ID'),client_secret:env('GOOGLE_CLIENT_SECRET'),redirect_uri:env('REDIRECT_URI'),grant_type:'authorization_code'});if(!t.access_token)return res.status(401).send('Falha na autenticação Google.');const u=await httpsGet('https://www.googleapis.com/oauth2/v2/userinfo',t.access_token);if(u.email!==env('ALLOWED_EMAIL'))return res.status(403).send(`<html><body style="font-family:sans-serif;text-align:center;padding:60px"><h2>Acesso negado</h2><p>Uso exclusivo do Prof. Felipe Vigneron Azevedo.</p><p style="color:#888">${u.email}</p></body></html>`);res.cookie('cla_session',makeToken(u.email),{httpOnly:true,secure:true,sameSite:'lax',maxAge:86400000*7,path:'/'});res.redirect('/')}catch(e){console.error(e);res.status(500).send('Erro interno.')}});
app.post('/cla-api',requireSession,upload.array('files'),async(req,res)=>{try{const pl=JSON.parse(req.body.payload||'{}'),mode=pl.mode,task=pl.task,inputs=pl.inputs||{},files=(req.files||[]).map(f=>({name:f.originalname,media_type:f.mimetype||'application/pdf',data:f.buffer.toString('base64')})),inp=inputs,fmt=(inp.formato||'').toUpperCase();inp.budget=(BUDGET[fmt]||{})[inp.nAulas||'']||{};if(mode==='A'){const sem=inp.semana||'X';if(task==='dpm_teorico')return res.json({docx:await gerarDPM(inp,files,'teorico'),filename:`DPM_Teorico_Sem${sem}.docx`,warnings:[]});if(task==='dpm_literario')return res.json({docx:await gerarDPM(inp,files,'literario'),filename:`DPM_Literario_Sem${sem}.docx`,warnings:[]});if(task==='quiz')return res.json({docx:await gerarQuizDoc(inp,files),filename:`Quiz_Sem${sem}.docx`,warnings:[]});if(task==='bimestral')return res.json({docx:await gerarBimestralDoc(inp,files),filename:`Bimestral_Sem${sem}.docx`,warnings:[]})}else if(mode==='B'){if(!inp.prompt)return res.status(400).json({error:'Campo prompt obrigatório.'});return res.json({text:await gerarOtimizador(inp),model:MODEL,warnings:[]})}else if(mode==='C'){if(task==='chat')return res.json({text:await gerarDevolutiva(inp,files||[]),model:MODEL,warnings:[]});return res.json({docx:await gerarDevolutivaDoc(inp,files||[]),filename:`Devolutiva_${inp.nivel||'trabalho'}.docx`,warnings:[]})}return res.status(400).json({error:`Modo desconhecido: ${mode}`})}catch(e){console.error(e);res.status(500).json({error:e.message||'Erro interno.'})}});
app.get('/',(req,res)=>{const p=path.join(__dirname,'index.html');console.log('Serving:',p);res.sendFile(p)});
app.use((req,res)=>res.status(404).send('Not found: '+req.url));
app.listen(PORT,()=>console.log(`Painel CLA rodando na porta ${PORT}`));
