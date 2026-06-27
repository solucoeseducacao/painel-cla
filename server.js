'use strict';
const express=require('express'),cookieParser=require('cookie-parser'),https=require('https'),crypto=require('crypto'),path=require('path'),Anthropic=require('@anthropic-ai/sdk'),{Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,AlignmentType,BorderStyle,WidthType,ShadingType}=require('docx');
const app=express(),PORT=process.env.PORT||3000;
function env(n){const v=process.env[n];if(!v)throw new Error(`Env missing: ${n}`);return v}
function sign(d){return crypto.createHmac('sha256',env('SESSION_SECRET')).update(d).digest('base64url')}
function makeToken(e){const p=Buffer.from(JSON.stringify({email:e,exp:Date.now()+86400000*7})).toString('base64url');return`${p}.${sign(p)}`}
function verifyToken(t){if(!t)return null;const i=t.indexOf('.');if(i<0)return null;const p=t.slice(0,i),s=t.slice(i+1);if(!p||!s)return null;try{const e=sign(p),eb=Buffer.from(e,'base64url'),sb=Buffer.from(s,'base64url');if(eb.length!==sb.length||!crypto.timingSafeEqual(eb,sb))return null;const d=JSON.parse(Buffer.from(p,'base64url').toString());if(d.exp<Date.now()||(d.email||'').toLowerCase()!==ALLOWED_EMAIL)return null;return d}catch{return null}}
function requireSession(req,res,next){const d=verifyToken(req.cookies?.cla_session);if(!d)return res.status(401).json({error:'Sessão inválida ou expirada.'});req.user=d;next()}
function httpsPost(url,data){return new Promise((resolve,reject)=>{const b=new URLSearchParams(data).toString(),u=new URL(url),r=https.request({hostname:u.hostname,path:u.pathname,method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(b)}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>{try{resolve(JSON.parse(d))}catch(e){reject(e)}})});r.on('error',reject);r.write(b);r.end()})}
function httpsGet(url,token){return new Promise((resolve,reject)=>{const u=new URL(url),r=https.request({hostname:u.hostname,path:u.pathname+u.search,method:'GET',headers:{Authorization:`Bearer ${token}`}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>{try{resolve(JSON.parse(d))}catch(e){reject(e)}})});r.on('error',reject);r.end()})}
const ALLOWED_EMAIL='felipevigneron@gmail.com'; // ÚNICO login permitido — fixo no código, não depende de env
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
const SYS_A=`IFF Campos Centro · Prof. Felipe Vigneron Azevedo · CLA. PT-BR. Tom: professor experiente, sem IA. Nunca inventar citações, páginas, datas ou autores — usar [a confirmar] ou omitir. Seguir estrutura pedida à risca.`;
const SYS_B=SYS_A+'\nModo: engenheiro de prompts. Preservar intenção do autor. Apontar a melhor alternativa e por quê.';
const SYS_C=SYS_A+'\nModo: orientador/banca. Devolutiva em 2ª pessoa, sem nota. Rigoroso, propositivo. Nomear problemas com termo exato.';
const GRUPOS_A_FMT_C='G4 Aplicação (Formato C): articular argumento teórico com o texto literário do DPM Literário desta semana.';
function inferirNivel(disc){if(!disc)return'intermediario';const d=disc.toLowerCase();if(d.includes('teoria liter')&&!d.includes('ii'))return'iniciante';if(d.includes('metodologia'))return'avancado';return'intermediario'}
function promptDPMTeorico(inp){const nd={iniciante:'iniciante (1º per.)',intermediario:'intermediária',avancado:'avançada/pós'}[inferirNivel(inp.disciplina)];const b=inp.budget,f=(inp.formato||'').toUpperCase(),ti=f==='B'?`leitura+discussão ${b.leit||'—'}min`:`conversa ${b.conv||'—'}min`,g4=f==='C'?'\nG4 Aplicação (Fmt C): articular com texto literário do DPM Literário desta semana.':'';return`**Disc:** ${inp.disciplina||''} | **Sem:** ${inp.semana||''} | **Aulas:** ${inp.nAulas||''} | **Fmt:** ${f} | ${nd}
**Tema:** ${inp.tema||''} | **Texto:** ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}
Tempos: ${ti} · grupo ${b.grp||'—'}min · saída ${b.cpd||'—'}min

Gere o DPM sem introduções. PT-BR. Nunca inventar citações/páginas — usar [a confirmar] ou omitir.

== VERSÃO ALUNOS ==
[Ref. ABNT antes das seções]

## S1 — Corrente Teórica
Tradição, problema central, método. 3–6 linhas.

## S2 — Tese Central
2–4 frases + ≥1 citação direta (SOBRENOME, ano, p. X).

## S3 — Conceitos-Chave
Tabela: Conceito (termo + pág.) | Explicação (2–3 frases). 3–7 itens, só do próprio texto.

## S4 — Parágrafos Centrais
3–6 citações diretas integrais com página. Não parafrasear.

## S5 — Perguntas de Grupo
G1 Tese: ≥3 pontos encadeados. G2 Mecanismo: recursos e função. G3 Tensão: onde hesita ou contradiz. G4 Aplicação: 2 conceitos do DPM → aula no EM.${g4} G5 Implicação: o que se segue para o campo ou prática docente.

== VERSÃO PROFESSOR ==
SP1 — Questão-Norteadora: oral antes do DPM, retomada ao final. Questão + resposta + páginas. ≠ perguntas de grupo.
SP2 — Tabela: Grupo | Resposta | Páginas. Omitir equívocos e refs. extras.
SP3 — Parágrafo do aluno: gere UMA pergunta discursiva única que contemple os pontos mais importantes do conteúdo trabalhado. Deve ser respondível em 3–5 frases (no máximo 2000 caracteres) — nunca exija resposta extensa, listas nem múltiplos itens. Formato EXATO:
## Parágrafo do aluno
[enunciado da pergunta, discursivo, sem alternativas]
**Resposta-modelo (3–5 frases, ≤2000 caracteres):** [resposta sintética com páginas]`}
function promptDPMLiterario(inp){const nd={iniciante:'iniciante',intermediario:'intermediária',avancado:'avançada'}[inferirNivel(inp.disciplina)];const b=inp.budget,f=(inp.formato||'').toUpperCase(),isM=(inp.disciplina||'').toLowerCase().includes('metodologia'),tipo=isM?'demonstrativo':'literário',g4=f==='C'?'G4 Aplicação: articular com DPM Teórico desta semana.':'G4 Intertexto: relações com outros textos evidenciadas pelo próprio texto.';return`**Disc:** ${inp.disciplina||''} | **Sem:** ${inp.semana||''} | **Fmt:** ${f} | ${nd}
**Tema:** ${inp.tema||''} | **Texto:** ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}
Tempos: discussão ${b.lit||b.leit||'—'}min · grupo ${b.grp||'—'}min · saída ${b.cpd||'—'}min

Gere o DPM sem introduções. PT-BR. Nunca inventar citações/páginas.

== VERSÃO ALUNOS ==
[Ref. ABNT antes das seções]

## S1 — Tese Central
2–4 frases com o argumento central do texto ${tipo}.

## S2 — Forma
Gênero · estrutura · narrador/voz · tempo · espaço · dicção${isM?' · argumento · metodologia demonstrada':''}.

## S3 — Conteúdo
Temas · personagens/agentes · conflito · desfecho.

## S4 — Contexto
Contexto histórico-literário${isM?'/acadêmico':''} · autor · período.

## S5 — Intertexto
Relações com outros textos evidenciadas pelo próprio texto.

## S6 — Parágrafos Centrais
3–5 citações diretas integrais com página. Não parafrasear.

## S7 — Perguntas de Grupo
G1 Forma · G2 Conteúdo · G3 Contexto · ${g4} · G5 Lacuna: o que o DPM não cobre.

== VERSÃO PROFESSOR ==
SP1 — Questão-Norteadora: oral antes do DPM. Questão + resposta + páginas.
SP2 — Tabela: Grupo | Resposta | Páginas. Omitir equívocos e refs. extras.
SP3 — Parágrafo do aluno: gere UMA pergunta discursiva única que contemple os pontos mais importantes do conteúdo trabalhado. Deve ser respondível em 3–5 frases (no máximo 2000 caracteres) — nunca exija resposta extensa, listas nem múltiplos itens. Formato EXATO:
## Parágrafo do aluno
[enunciado da pergunta, discursivo, sem alternativas]
**Resposta-modelo (3–5 frases, ≤2000 caracteres):** [resposta sintética com páginas]`}
function promptQuiz(inp){const b=inp.budget;return`Quiz — ${inp.disciplina||''} | Sem ${inp.semana||''} | ${inp.tema||''} | ${inp.referencias||''}

5 questões A–D, sem cabeçalho. Formato: nº + enunciado → A/B/C/D → Gabarito: [letra] → linha em branco.
Q1: reformulação da questão-norteadora. Q2–5: baseadas no texto, sem repetir perguntas de grupo.
Após Q5: linha divisória + COMENTÁRIOS PROFESSOR no formato: 1 - "comentário" / 2 - "comentário" / 3 - "comentário" / 4 - "comentário" / 5 - "comentário" (incluir SOBRENOME, ano, p. X em cada comentário).`}
function promptBimestral(inp){return`Bimestral — ${inp.disciplina||''} | Sem ${inp.semana||''} | ${inp.referencias||''}

2 questões A–D, sem cabeçalho. Q1 (compreensão) · Q2 (interpretação). Formato: enunciado → A/B/C/D → linha em branco.
Após Q2: linha divisória + GABARITO PROFESSOR: nº · resposta · comentário · (SOBRENOME, ano, p. X).`}
function toFileContent(files){const arr=(files||[]).map(f=>{
  // Normalizar media_type para tipos aceitos pela Anthropic
  let mt=f.media_type||'application/pdf';
  if(mt==='application/octet-stream'||mt==='application/msword'){
    if(f.name&&f.name.match(/\.docx?$/i))mt='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  }
  return{type:'document',source:{type:'base64',media_type:mt,data:f.data}};
});if(arr.length)arr[arr.length-1].cache_control={type:'ephemeral'};return arr;}
function txt(res){const b=((res&&res.content)||[]).find(x=>x.type==='text');return b?b.text:''}
async function gerarDPM(inp,files,tipo){const prompt=tipo==='teorico'?promptDPMTeorico(inp):promptDPMLiterario(inp);const res=await client.messages.create({model:MODEL,max_tokens:3600,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:prompt}]}]});const text=txt(res),si=text.search(/==\s*VERS[ÃA]O\s+PROFESSOR\s*==|(?:^|\n)\s*SP1\s*[—\-:]/i),aR=si>0?text.slice(0,si):text,pR=si>0?text.slice(si):'',aT=aR.replace(/==\s*VERS[ÃA]O\s+ALUNOS\s*==/i,'').trim(),pT=pR.replace(/==\s*VERS[ÃA]O\s+PROFESSOR\s*==/i,'').trim(),isM=(inp.disciplina||'').toLowerCase().includes('metodologia'),docName=tipo==='teorico'?'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Teórico':isM?'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Demonstrativo':'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Literário',label=tipo==='teorico'?'DPM TEÓRICO':isM?'DPM DEMONSTRATIVO':'DPM LITERÁRIO';const children=[makeHeader(inp.disciplina,inp.semana,docName),p(''),p(`${label} — VERSÃO ALUNOS`,{bold:true,size:28,sb:200}),p(''),...mdToDocx(aT)];if(pT)children.push(...separador(),faixaConf(),p(''),...mdToDocx(pT));return Packer.toBase64String(makeDoc(children))}
async function gerarQuizDoc(inp,files){const res=await client.messages.create({model:MODEL_FAST,max_tokens:1200,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptQuiz(inp)}]}]});return Packer.toBase64String(makeDoc([...mdToDocx(txt(res))]))}
async function gerarBimestralDoc(inp,files){const res=await client.messages.create({model:MODEL_FAST,max_tokens:800,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptBimestral(inp)}]}]});return Packer.toBase64String(makeDoc([...mdToDocx(txt(res))]))}
async function gerarOtimizador(inp){const res=await client.messages.create({model:MODEL,max_tokens:1500,system:SYS_B,messages:[{role:'user',content:`${inp.origem==='painel'?'Material do Painel CLA.\n':''}Queixa/objetivo: ${inp.queixa||'não especificado'}\n\n${inp.prompt}\n\nEntregar: 1) DIAGNÓSTICO 2) VERSÃO OTIMIZADA 3) O QUE MUDOU E POR QUÊ`}]});return txt(res)}
async function gerarDevolutiva(inp,files){const papel=inp.papel==='banca'?'banca (avaliativa)':'orientador (formativa)',fase={inicio:'início',andamento:'andamento',concluido:'concluído (pré-banca)'}[inp.fase]||'',nivel={artigo:'artigo/TCC',dissertacao:'dissertação',tese:'tese'}[inp.nivel]||inp.nivel;const res=await client.messages.create({model:MODEL,max_tokens:2500,system:SYS_C,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:`Papel: ${papel}${inp.papel!=='banca'?' · Fase: '+fase:''} · Nível: ${nivel}\n${inp.foco?'Foco: '+inp.foco:''}\n${inp.contexto?'Contexto/trechos:\n'+inp.contexto:''}\n${files&&files.length?'Trabalho anexado acima.':'Usar contexto/trechos fornecidos.'}\n\nCritérios (por peso): 1) Cumprimento dos objetivos 2) Originalidade 3) Fundamentação teórica 4) Correção conceitual · Clareza · Consistência · ABNT\n\nEstrutura:\n1 — Leitura geral (2–4 frases)\n2 — Pontos por critério (só onde há algo a dizer)\n3 — Apontamentos cirúrgicos: trecho → problema → sugestão → fonte\n4 — Prioridades (2–3 providências)\n5 — Próximo passo`}]}]});return txt(res)}
async function gerarDevolutivaDoc(inp,files){const text=await gerarDevolutiva(inp,files),papel=inp.papel==='banca'?'Banca':'Orientador',fase={inicio:'Início',andamento:'Andamento',concluido:'Concluído'}[inp.fase]||'',nivel={artigo:'Artigo/TCC',dissertacao:'Dissertação',tese:'Tese'}[inp.nivel]||inp.nivel;return Packer.toBase64String(makeDoc([p(`Devolutiva · ${papel}${fase?' · '+fase:''} · ${nivel}`,{size:18,color:'555555',sb:0}),p(''),...mdToDocx(text)]))}
function promptTabela(inp,papeis,nq){const f=(inp.formato||'').toUpperCase();return `Folha de Resposta DPM — ${inp.disciplina||''} | Sem ${inp.semana||''} | Formato ${f}
Texto(s) da semana: ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}

Gere ${nq} pergunta(s) norteadora(s) por grupo, ancorada(s) no papel do grupo e no(s) texto(s) da semana. PT-BR. Sem respostas, sem numerar alternativas. Nunca inventar citações/páginas.
Formato EXATO de saída, uma linha por grupo: "G<n>: <pergunta>"${nq>1?' — para os dois itens do mesmo grupo, separe-os por " ||| ".':'.'}
Grupos e papéis:
${papeis.map((pp,i)=>`G${i+1} (${pp})`).join('\n')}`;}
function parseTabela(text,n){const map={};(text||'').split('\n').forEach(l=>{const m=l.match(/^\s*G\s*(\d+)\s*[:\-—]\s*(.+)$/);if(m)map[+m[1]]=m[2].trim();});const out=[];for(let i=1;i<=n;i++)out.push((map[i]||'').split('|||').map(s=>s.trim()).filter(Boolean));return out;}
async function gerarTabelaDoc(inp,files){const f=(inp.formato||'').toUpperCase();const PAPEIS=({A:['Tese','Contexto','Contradição','Aplicação','Implicação'],B:['Forma','Conteúdo','Contradição','Transposição','Aplicação'],C:['Tese·Teórico + Forma·Literário','Contexto·Teórico + Conteúdo·Literário','Contradição·Misto','Aplicação·Teórico + Transposição·Literário','Implicação·Teórico + Aplicação·Literário']})[f]||['Tese','Contexto','Contradição','Aplicação','Implicação'];const nq=f==='C'?2:1;const res=await client.messages.create({model:MODEL,max_tokens:1500,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptTabela(inp,PAPEIS,nq)}]}]});const perguntas=parseTabela(txt(res),PAPEIS.length);const fl=({A:'Teórico',B:'Literário',C:'Misto'})[f]||f;const children=[makeHeader(inp.disciplina,inp.semana,`DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM ${fl}`),p(''),p(`FOLHA DE RESPOSTA DPM — ${fl.toUpperCase()}`,{bold:true,size:28,sb:200}),p('Preencha o parágrafo do seu grupo no espaço indicado.',{italic:true,color:'555555',sb:0,sa:120})];PAPEIS.forEach((papel,i)=>{children.push(p(`Grupo ${i+1} — ${papel}`,{bold:true,size:26,color:'2E6B3E',sb:240,sa:60}));const qs=(perguntas[i]&&perguntas[i].length)?perguntas[i]:['[pergunta a confirmar]'];qs.forEach((q,qi)=>children.push(p((nq>1?(qi+1)+'. ':'')+q,{sb:30,sa:50})));children.push(p('Parágrafo do grupo:',{italic:true,size:22,color:'555555',sb:60,sa:60}));for(let k=0;k<8;k++)children.push(p('______________________________________________________________',{sb:50,sa:50,color:'BBBBBB'}));});return Packer.toBase64String(makeDoc(children));}
app.use(cookieParser());app.use(express.json({limit:'50mb'}));
app.use((req,res,next)=>{const o=req.headers.origin,a=process.env.SITE_URL||o;if(a){res.setHeader('Access-Control-Allow-Origin',a);res.setHeader('Vary','Origin');}res.setHeader('Access-Control-Allow-Credentials','true');res.setHeader('Access-Control-Allow-Methods','GET, POST, OPTIONS');res.setHeader('Access-Control-Allow-Headers','Content-Type');if(req.method==='OPTIONS')return res.sendStatus(200);next()});
app.get('/auth-check',(req,res)=>{const d=verifyToken(req.cookies?.cla_session);if(d)return res.json({ok:true,email:d.email,authUrl:null});const p=new URLSearchParams({client_id:env('GOOGLE_CLIENT_ID'),redirect_uri:env('REDIRECT_URI'),response_type:'code',scope:'openid email profile',prompt:'select_account'});res.json({ok:false,email:null,authUrl:`https://accounts.google.com/o/oauth2/v2/auth?${p.toString()}`})});
app.get('/auth-callback',async(req,res)=>{const c=req.query.code;if(!c)return res.status(400).send('Código ausente.');try{const t=await httpsPost('https://oauth2.googleapis.com/token',{code:c,client_id:env('GOOGLE_CLIENT_ID'),client_secret:env('GOOGLE_CLIENT_SECRET'),redirect_uri:env('REDIRECT_URI'),grant_type:'authorization_code'});if(!t.access_token)return res.status(401).send('Falha na autenticação Google.');const u=await httpsGet('https://www.googleapis.com/oauth2/v2/userinfo',t.access_token);if((u.email||'').toLowerCase()!==ALLOWED_EMAIL)return res.status(403).send(`<html><body style="font-family:sans-serif;text-align:center;padding:60px"><h2>Acesso negado</h2><p>Uso exclusivo do Prof. Felipe Vigneron Azevedo.</p><p style="color:#888">${u.email}</p></body></html>`);res.cookie('cla_session',makeToken(u.email),{httpOnly:true,secure:true,sameSite:'lax',maxAge:86400000*7,path:'/'});res.redirect('/')}catch(e){console.error(e);res.status(500).send('Erro interno.')}});
app.post('/cla-api',requireSession,async(req,res)=>{try{const {mode,task,inputs={},files=[]}=req.body||{},inp=inputs,fmt=(inp.formato||'').toUpperCase();inp.budget=(BUDGET[fmt]||{})[inp.nAulas||'']||{};if(mode==='A'){const sem=inp.semana||'X';if(task==='dpm_teorico')return res.json({docx:await gerarDPM(inp,files,'teorico'),filename:`DPM_Teorico_Sem${sem}.docx`,warnings:[]});if(task==='dpm_literario')return res.json({docx:await gerarDPM(inp,files,'literario'),filename:`DPM_Literario_Sem${sem}.docx`,warnings:[]});if(task==='quiz')return res.json({docx:await gerarQuizDoc(inp,files),filename:`Quiz_Sem${sem}.docx`,warnings:[]});if(task==='bimestral')return res.json({docx:await gerarBimestralDoc(inp,files),filename:`Bimestral_Sem${sem}.docx`,warnings:[]});if(task==='tabela_dpm')return res.json({docx:await gerarTabelaDoc(inp,files),filename:`FolhaResposta_DPM_Sem${sem}.docx`,warnings:[]})}else if(mode==='B'){if(!inp.prompt)return res.status(400).json({error:'Campo prompt obrigatório.'});return res.json({text:await gerarOtimizador(inp),model:MODEL,warnings:[]})}else if(mode==='C'){if(task==='chat')return res.json({text:await gerarDevolutiva(inp,files),warnings:[]});if(task==='docx')return res.json({docx:await gerarDevolutivaDoc(inp,files),filename:'Devolutiva.docx',warnings:[]})}res.status(400).json({error:'Modo/tarefa não reconhecido.'})}catch(e){console.error(e);res.status(500).json({error:e.message||'Erro interno.'})}});
app.get('/',(req,res)=>res.sendFile(path.join(__dirname,'index.html')));
app.listen(PORT,()=>console.log(`CLA running on ${PORT}`));
