'use strict';
const express=require('express'),cookieParser=require('cookie-parser'),https=require('https'),crypto=require('crypto'),path=require('path'),Anthropic=require('@anthropic-ai/sdk'),{Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,AlignmentType,BorderStyle,WidthType,ShadingType,PageBreak}=require('docx');
const app=express(),PORT=process.env.PORT||3000;
function env(n){const v=process.env[n];if(!v)throw new Error(`Env missing: ${n}`);return v}
function sign(d){return crypto.createHmac('sha256',env('SESSION_SECRET')).update(d).digest('base64url')}
function makeToken(e){const p=Buffer.from(JSON.stringify({email:e,exp:Date.now()+86400000*7})).toString('base64url');return`${p}.${sign(p)}`}
function verifyToken(t){if(!t)return null;const i=t.indexOf('.');if(i<0)return null;const p=t.slice(0,i),s=t.slice(i+1);if(!p||!s)return null;try{const e=sign(p),eb=Buffer.from(e,'base64url'),sb=Buffer.from(s,'base64url');if(eb.length!==sb.length||!crypto.timingSafeEqual(eb,sb))return null;const d=JSON.parse(Buffer.from(p,'base64url').toString());if(d.exp<Date.now()||(d.email||'').toLowerCase()!==ALLOWED_EMAIL)return null;return d}catch{return null}}
function requireSession(req,res,next){const d=verifyToken(req.cookies?.cla_session);if(!d)return res.status(401).json({error:'Sessão inválida ou expirada.'});req.user=d;next()}
function httpsPost(url,data){return new Promise((resolve,reject)=>{const b=new URLSearchParams(data).toString(),u=new URL(url),r=https.request({hostname:u.hostname,path:u.pathname,method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(b)}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>{try{resolve(JSON.parse(d))}catch(e){reject(e)}})});r.on('error',reject);r.write(b);r.end()})}
function httpsGet(url,token){return new Promise((resolve,reject)=>{const u=new URL(url),r=https.request({hostname:u.hostname,path:u.pathname+u.search,method:'GET',headers:{Authorization:`Bearer ${token}`}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>{try{resolve(JSON.parse(d))}catch(e){reject(e)}})});r.on('error',reject);r.end()})}
const ALLOWED_EMAIL='felipevigneron@gmail.com'; // ÚNICO login permitido — fixo no código, não depende de env
const client=new Anthropic({apiKey:process.env.ANTHROPIC_API_KEY}),MODEL='claude-sonnet-4-6',MODEL_FAST='claude-haiku-4-5',MODEL_OPUS='claude-opus-4-8';
// ── SELEÇÃO REAL DE MODELO + ESFORÇO (espelha o server do Super App) ───────────
// Resolve a chave do front ('haiku'|'sonnet'|'opus') para o ID; usa o padrão por
// tipo de geração quando vier vazio. Padrão econômico = Sonnet (Opus só sob demanda).
function _modeloDe(chave,padrao){if(chave==='haiku')return MODEL_FAST;if(chave==='sonnet')return MODEL;if(chave==='opus')return MODEL_OPUS;return padrao||MODEL;}
// Esforço: SÓ Sonnet/Opus aceitam output_config.effort — no Haiku 4.5 a API dá 400.
const _ESFORCO={baixo:'low',medio:'medium','médio':'medium',alto:'high',low:'low',medium:'medium',high:'high'};
// Devolve {model[,thinking,output_config]} a espalhar no corpo da chamada.
function _iaExtra(inp,padrao){const m=_modeloDe(inp&&inp.modelo,padrao);const eff=(m!==MODEL_FAST)?_ESFORCO[String(inp&&inp.esforco||'').toLowerCase()]:null;return eff?{model:m,thinking:{type:'adaptive',display:'omitted'},output_config:{effort:eff}}:{model:m};}
// ── Custo, rate-limit e captura de uso (observabilidade de custo/cache) ─────────
// Preço US$/MTok: Haiku 1/5 · Sonnet 3/15 · Opus 5/25. Cache: escrita 1,25× · leitura 0,1×.
function _precoMTok(m){return m===MODEL_OPUS?{i:5,o:25}:m===MODEL_FAST?{i:1,o:5}:{i:3,o:15};}
function _custoIA(m,u){u=u||{};const p=_precoMTok(m);const i=u.input_tokens||0,o=u.output_tokens||0,cw=u.cache_creation_input_tokens||0,cr=u.cache_read_input_tokens||0;return (i*p.i+o*p.o+cw*p.i*1.25+cr*p.i*0.1)/1e6;}
function _nomeModelo(m){return m===MODEL_OPUS?'Opus 4.8':m===MODEL_FAST?'Haiku 4.5':'Sonnet 4.6';}
// Rate-limit IA: 40 chamadas/min por usuário (espelha o Super App).
const _iaRate=new Map();
function iaRateOk(email){const now=Date.now(),k=email||'anon',calls=(_iaRate.get(k)||[]).filter(t=>now-t<60000);if(calls.length>=40)return false;calls.push(now);_iaRate.set(k,calls);return true;}
// Último uso de tokens (preenchido pelo wrapper) → exposto no /cla-api p/ o cliente ver custo/cache.
let _ultimoUso=null;
// C13: registra o uso das gerações do CLA no contador central do Super App (GAS __uso_api__).
// GAS_URL é público (mesmo da aba do aluno); GAS_SENHA é segredo (env). Sem a env → no-op silencioso.
const GAS_URL_USO = process.env.GAS_URL || 'https://script.google.com/macros/s/AKfycby_pu_oDSP2nJNWIooa1wEl-fVxxvp8KyV_KNbS2ogPcWDshzxYCXSx5v6KtBxztarRxg/exec';
function _registrarUsoGAS(u){ try{ if(!process.env.GAS_SENHA||!u) return; fetch(GAS_URL_USO,{method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify({acao:'registrarUsoApi',senha:process.env.GAS_SENHA,modelo:u.modelo,in:u.in,out:u.out,custo:u.custo})}).catch(function(){}); }catch(_){} }
// Prova de prompt-cache nos logs do Render: cacheR>0 ⇒ PDFs reaproveitados do cache.
(function(){const _create=client.messages.create.bind(client.messages);client.messages.create=async function(opts){const res=await _create(opts);try{const u=(res&&res.usage)||{},m=opts&&opts.model;_ultimoUso={modelo:_nomeModelo(m),modeloId:m,in:u.input_tokens||0,out:u.output_tokens||0,cacheW:u.cache_creation_input_tokens||0,cacheR:u.cache_read_input_tokens||0,custo:Number(_custoIA(m,u).toFixed(4))};console.log('[uso]',m,JSON.stringify(_ultimoUso));_registrarUsoGAS(_ultimoUso);}catch(_){}return res;};})();
const BUDGET={A:{'2':{conv:20,quiz:10,entrada:10,grp:20,reg:15,cpd:15},'3':{conv:30,quiz:10,entrada:10,grp:35,reg:20,cpd:15},'4':{conv:45,quiz:10,entrada:15,grp:50,reg:25,cpd:15}},B:{'2':{leit:30,quiz:10,grp:20,reg:15,cpd:15},'3':{leit:45,quiz:10,grp:35,reg:15,cpd:15},'4':{leit:65,quiz:10,grp:50,reg:20,cpd:15}},C:{'2':{conv:15,lit:15,quiz:10,grp:20,reg:15,cpd:15},'3':{conv:20,lit:20,quiz:10,grp:35,reg:20,cpd:15},'4':{conv:30,lit:30,quiz:10,grp:45,reg:25,cpd:20}}};
const A4={size:{width:11906,height:16838},margin:{top:1440,right:1440,bottom:1440,left:1440}};
const BDR={style:BorderStyle.SINGLE,size:1,color:'999999'},BORDERS={top:BDR,bottom:BDR,left:BDR,right:BDR};
function tc(children,opts={}){return new TableCell({borders:BORDERS,shading:{fill:opts.fill||'FFFFFF',type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},width:{size:opts.w||4513,type:WidthType.DXA},children})}
function makeHeader(disc,semana,doc){const W1=5416,W2=3610;return new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[W1,W2],rows:[new TableRow({children:[tc([new Paragraph({children:[new TextRun({text:'Licenciatura em Letras — Português e Literaturas',bold:true,size:22,font:'Arial'})]}),new Paragraph({children:[new TextRun({text:`Disciplina: ${disc||''}`,size:21,font:'Arial'})]}),new Paragraph({children:[new TextRun({text:`Professor: Felipe Vigneron Azevedo  |  Semana: ${semana||''}`,size:21,font:'Arial'})]})],{fill:'B3E5A0',w:W1}),tc([new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:doc||'',bold:true,size:22,font:'Arial'})]})],{fill:'FFFFFF',w:W2})]}),new TableRow({children:[tc([new Paragraph({children:[new TextRun({text:'Nome do(a) Estudante: _______________________________________________',size:22,font:'Arial'})]})],{w:W1}),tc([new Paragraph({children:[new TextRun({text:'Data: ___/___/______',size:22,font:'Arial'})]})],{w:W2})]})]})}
function faixaConf(txt){const l=txt||'VERSÃO PROFESSOR — CONFIDENCIAL — NÃO DISTRIBUIR AOS ALUNOS';return new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],rows:[new TableRow({children:[tc([new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:l,bold:true,color:'FFFFFF',size:24,font:'Arial'})]})],{fill:'8B0000',w:9026})]})]})}
function separador(){return[p('',{sb:400,sa:0,borderTop:true,borderColor:'8B0000'}),p('✂  SEPARAR AQUI — A partir daqui: DPM PROFESSOR (CONFIDENCIAL)',{bold:true,color:'8B0000',center:true,sb:80,sa:80}),p('',{sb:0,sa:400,borderTop:true,borderColor:'8B0000'})]}
function p(text,opts={}){return new Paragraph({alignment:opts.center?AlignmentType.CENTER:AlignmentType.LEFT,spacing:{before:opts.sb??120,after:opts.sa??80},border:opts.borderTop?{top:{style:BorderStyle.SINGLE,size:6,color:opts.borderColor||'000000'}}:undefined,children:[new TextRun({text:text||'',bold:!!opts.bold,italic:!!opts.italic,size:opts.size||24,font:'Arial',color:opts.color||'000000',underline:opts.underline?{}:undefined})]})}
function parseInlineRuns(text){const r=[],rx=/\*\*(.+?)\*\*|\*(.+?)\*|([^*]+)/g;let m;while((m=rx.exec(text))!==null){if(m[1])r.push(new TextRun({text:m[1],bold:true,size:24,font:'Arial'}));else if(m[2])r.push(new TextRun({text:m[2],italic:true,size:24,font:'Arial'}));else if(m[3])r.push(new TextRun({text:m[3],size:24,font:'Arial'}))}return r.length?r:[new TextRun({text,size:24,font:'Arial'})]}
function makeTableFromMd(rows){const colC=rows[0].length,colW=Math.floor(9026/colC);return new Table({width:{size:9026,type:WidthType.DXA},columnWidths:Array(colC).fill(colW),rows:rows.map((row,ri)=>new TableRow({children:row.map(cell=>new TableCell({borders:ri===0?{top:{style:BorderStyle.SINGLE,size:4,color:'2E6B3E'},bottom:{style:BorderStyle.SINGLE,size:4,color:'2E6B3E'},left:{style:BorderStyle.SINGLE,size:2,color:'999999'},right:{style:BorderStyle.SINGLE,size:2,color:'999999'}}:{top:{style:BorderStyle.SINGLE,size:2,color:'999999'},bottom:{style:BorderStyle.SINGLE,size:2,color:'999999'},left:{style:BorderStyle.SINGLE,size:2,color:'999999'},right:{style:BorderStyle.SINGLE,size:2,color:'999999'}},shading:{fill:ri===0?'E8F5E9':'FFFFFF',type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},width:{size:colW,type:WidthType.DXA},children:[new Paragraph({spacing:{before:60,after:60},children:ri===0?[new TextRun({text:cell.trim(),bold:true,size:23,font:'Arial',color:'1B5E20'})]:parseInlineRuns(cell.trim())})]}))}))})}
function mdToDocx(text){if(!text)return[];const out=[],lines=text.split('\n');let i=0;while(i<lines.length){const line=lines[i];if(/^-{3,}$/.test(line.trim())){i++;continue}if(!line.trim()){out.push(p(''));i++;continue}if(line.startsWith('### ')){out.push(p(line.slice(4),{bold:true,size:28,sb:200,sa:80}));i++;continue}if(line.startsWith('## ')){out.push(p(line.slice(3),{bold:true,size:26,sb:240,sa:100,color:'2E6B3E'}));i++;continue}if(line.startsWith('# ')){out.push(p(line.slice(2),{bold:true,size:30,sb:280,sa:120}));i++;continue}if(line.startsWith('|')){const tbl=[];while(i<lines.length&&lines[i].startsWith('|')){if(!/^\|[-:| ]+\|$/.test(lines[i])){const cells=lines[i].split('|').slice(1,-1);tbl.push(cells)}i++}if(tbl.length>0){out.push(p(''));out.push(makeTableFromMd(tbl));out.push(p(''))}continue}const indent=line.startsWith('- '),rest=indent?line.slice(2):line;out.push(new Paragraph({spacing:{before:80,after:80},indent:indent?{left:360}:undefined,children:parseInlineRuns(rest)}));i++}return out}
function makeDoc(children){return new Document({styles:{default:{document:{run:{font:'Arial',size:24}}}},sections:[{properties:{page:A4},children}]})}
const SYS_A=`IFF Campos Centro · Prof. Felipe Vigneron Azevedo · CLA. PT-BR. Tom: professor experiente, sem IA.
REGRA Nº 1 — VERDADE ACIMA DE TUDO: não invente, não minta, não delire, não crie nada. Aja como um cientista de alta performance obcecado por exatidão. TODO o material precisa ser verificado nos arquivos anexados.
CITAÇÕES E TRECHOS: cite/transcreva APENAS o que está literalmente nos arquivos anexados — o trecho citado deve existir no arquivo (não parafrasear como se fosse citação direta).
PÁGINAS: use o número de página IMPRESSO na obra digitalizada (não o número da página do leitor de PDF). Se a obra não trouxer página impressa, NÃO invente: oriente-se pela página do leitor de PDF e escreva exatamente "p. X, conforme arquivo". Só use "[a confirmar]" quando for realmente impossível localizar a informação no anexo. Nunca atribua autor, ano ou página que você não confirmou no arquivo.
Seguir a estrutura pedida à risca.`;
// Bloco de citação reaproveitado nos prompts (reforça a REGRA Nº 1 dentro da mensagem do usuário).
const REGRA_CIT=`REGRA Nº 1 — VERDADE: não inventar, não mentir, não delirar, não criar nada; verificar TUDO nos arquivos anexados, como um cientista obcecado por exatidão.
- Toda citação/trecho deve existir LITERALMENTE no(s) arquivo(s) anexado(s).
- Página = número impresso na obra digitalizada (não a página do leitor de PDF). Sem página impressa: escreva "p. X, conforme arquivo" (página do leitor). Só "[a confirmar]" se for impossível localizar — nunca inventar.`;
const SYS_B=SYS_A+'\nModo: engenheiro de prompts. Preservar intenção do autor. Apontar a melhor alternativa e por quê.';
const SYS_C=SYS_A+'\nModo: orientador/banca. Devolutiva em 2ª pessoa, sem nota. Rigoroso, propositivo. Nomear problemas com termo exato.';
const GRUPOS_A_FMT_C='G4 Aplicação (Formato C): articular argumento teórico com o texto literário do DPM Literário desta semana.';
function inferirNivel(disc){if(!disc)return'intermediario';const d=disc.toLowerCase();if(d.includes('teoria liter')&&!d.includes('ii'))return'iniciante';if(d.includes('metodologia'))return'avancado';return'intermediario'}
function promptDPMTeorico(inp){const nd={iniciante:'iniciante (1º per.)',intermediario:'intermediária',avancado:'avançada/pós'}[inferirNivel(inp.disciplina)];const b=inp.budget,f=(inp.formato||'').toUpperCase(),ti=f==='B'?`leitura+discussão ${b.leit||'—'}min`:`conversa ${b.conv||'—'}min`,g4=f==='C'?'\nG4 Aplicação (Fmt C): articular com texto literário do DPM Literário desta semana.':'';return`**Disc:** ${inp.disciplina||''} | **Sem:** ${inp.semana||''} | **Aulas:** ${inp.nAulas||''} | **Fmt:** ${f} | ${nd}
**Tema:** ${inp.tema||''} | **Texto:** ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}
Tempos: ${ti} · grupo ${b.grp||'—'}min · saída ${b.cpd||'—'}min

Gere o DPM sem introduções. PT-BR.
${REGRA_CIT}

Use os títulos de seção EXATAMENTE como abaixo (sem códigos como "S1", "S2", "SP1" — apenas o nome).

== VERSÃO ALUNOS ==
[Ref. ABNT antes das seções]

## Corrente Teórica
Tradição, problema central, método. 3–6 linhas.

## Tese Central
2–4 frases + ≥1 citação direta entre aspas com (SOBRENOME, ano, p. X) extraída do arquivo.

## Conceitos-Chave
Tabela: Conceito (termo + pág. impressa do arquivo) | Explicação (2–3 frases). 3–7 itens, só do próprio texto. Cada conceito com a página onde aparece no arquivo (ou "p. X, conforme arquivo").

## Parágrafos Centrais
3–6 citações diretas integrais entre aspas, cada uma com (SOBRENOME, ano, p. X) do arquivo. Transcrever literalmente do anexo — não parafrasear.

## Perguntas de Grupo
Grupo 1 — Tese: ≥3 pontos encadeados. Grupo 2 — Mecanismo: recursos e função. Grupo 3 — Tensão: onde hesita ou contradiz. Grupo 4 — Aplicação: 2 conceitos do DPM → aula no EM.${g4} Grupo 5 — Implicação: o que se segue para o campo ou prática docente.

== VERSÃO PROFESSOR ==
(Obrigatório: TODA questão desta versão vem com gabarito/sugestão de resposta ancorado no arquivo, com páginas conforme a REGRA Nº 1. Não escreva esta linha no documento.)

## Questão-Norteadora
UMA pergunta para debate ORAL antes da leitura, retomada ao final (≠ perguntas de grupo). Enunciado COMPLETO + resposta-guia para o professor conduzir o debate (com páginas).

## Gabarito das Perguntas de Grupo
Para CADA grupo (Tese, Mecanismo, Tensão, Aplicação, Implicação): repita a pergunta + resposta-modelo concisa ancorada no texto, com página(s). Nenhum grupo sem gabarito. Com mais de uma obra anexada, atribua cada citação à obra correta.

## Parágrafo do Aluno
UMA pergunta discursiva única que contemple os pontos mais importantes do conteúdo. Respondível em 3–5 frases (máx. 2000 caracteres) — nunca exija resposta extensa, listas nem múltiplos itens. Enunciado completo (discursivo, sem alternativas), seguido de:
**Resposta-modelo (3–5 frases, ≤2000 caracteres):** [resposta sintética com páginas]`}
function promptDPMLiterario(inp){const nd={iniciante:'iniciante',intermediario:'intermediária',avancado:'avançada'}[inferirNivel(inp.disciplina)];const b=inp.budget,f=(inp.formato||'').toUpperCase(),isM=(inp.disciplina||'').toLowerCase().includes('metodologia'),tipo=isM?'demonstrativo':'literário',g4=f==='C'?'G4 Aplicação: articular com DPM Teórico desta semana.':'G4 Intertexto: relações com outros textos evidenciadas pelo próprio texto.';return`**Disc:** ${inp.disciplina||''} | **Sem:** ${inp.semana||''} | **Fmt:** ${f} | ${nd}
**Tema:** ${inp.tema||''} | **Texto:** ${inp.referencias||''}${inp.obs?'\nObs.: '+inp.obs:''}
Tempos: discussão ${b.lit||b.leit||'—'}min · grupo ${b.grp||'—'}min · saída ${b.cpd||'—'}min

Gere o DPM sem introduções. PT-BR.
${REGRA_CIT}

Use os títulos de seção EXATAMENTE como abaixo (sem códigos como "S1", "S2", "SP1" — apenas o nome).

== VERSÃO ALUNOS ==
[Ref. ABNT antes das seções]

## Tese Central
2–4 frases com o argumento central do texto ${tipo} + ≥1 citação direta entre aspas com (SOBRENOME, ano, p. X) do arquivo.

## Forma
Gênero · estrutura · narrador/voz · tempo · espaço · dicção${isM?' · argumento · metodologia demonstrada':''}.

## Conteúdo
Temas · personagens/agentes · conflito · desfecho.

## Contexto
Contexto histórico-literário${isM?'/acadêmico':''} · autor · período.

## Intertexto
Relações com outros textos evidenciadas pelo próprio texto.

## Parágrafos Centrais
3–5 citações diretas integrais entre aspas, cada uma com (SOBRENOME, ano, p. X) do arquivo. Transcrever literalmente do anexo — não parafrasear.

## Perguntas de Grupo
Grupo 1 — Forma · Grupo 2 — Conteúdo · Grupo 3 — Contexto · Grupo 4 — ${g4} · Grupo 5 — Lacuna: o que o DPM não cobre.

== VERSÃO PROFESSOR ==
(Obrigatório: TODA questão desta versão vem com gabarito/sugestão de resposta ancorado no arquivo, com páginas conforme a REGRA Nº 1. Não escreva esta linha no documento.)

## Questão-Norteadora
UMA pergunta para debate ORAL antes da leitura, retomada ao final (≠ perguntas de grupo). Enunciado COMPLETO + resposta-guia para o professor (com páginas).

## Gabarito das Perguntas de Grupo
Para CADA grupo (Forma, Conteúdo, Contexto, Grupo 4, Lacuna): repita a pergunta + resposta-modelo concisa ancorada no texto, com página(s). Nenhum grupo sem gabarito. Com mais de uma obra anexada, atribua cada citação à obra correta.

## Parágrafo do Aluno
UMA pergunta discursiva única que contemple os pontos mais importantes do conteúdo. Respondível em 3–5 frases (máx. 2000 caracteres) — nunca exija resposta extensa, listas nem múltiplos itens. Enunciado completo (discursivo, sem alternativas), seguido de:
**Resposta-modelo (3–5 frases, ≤2000 caracteres):** [resposta sintética com páginas]`}
function promptQuiz(inp){const b=inp.budget;return`Quiz — ${inp.disciplina||''} | Sem ${inp.semana||''} | ${inp.tema||''} | ${inp.referencias||''}
${REGRA_CIT}

5 questões de múltipla escolha A–D, sem cabeçalho. Para CADA questão, nesta ordem:
nº + enunciado
A) ... / B) ... / C) ... / D) ...
Gabarito: [APENAS a letra — SEM comentário]
(linha em branco)
Q1 = reformulação da questão-norteadora gerada no DPM desta semana. Q2–Q5 baseadas no(s) texto(s), sem repetir as perguntas de grupo.
Depois de TODAS as questões, escreva numa linha isolada EXATAMENTE: %%PAGEBREAK%%
Em seguida: o título "## Gabarito comentado" e, para cada questão, UM comentário SUCINTO (1 linha) no formato: "1. B — <comentário curto> (SOBRENOME, ano, p. X)". Comentários enxutos.`}
function promptBimestral(inp){return`Bimestral — ${inp.disciplina||''} | Sem ${inp.semana||''} | ${inp.referencias||''}
${REGRA_CIT}

2 questões de múltipla escolha A–D, sem cabeçalho. Q1 (compreensão) · Q2 (interpretação). Para CADA questão, nesta ordem:
enunciado
A) ... / B) ... / C) ... / D) ...
Gabarito: [APENAS a letra — SEM justificativa]
(linha em branco)
Depois das 2 questões, escreva numa linha isolada EXATAMENTE: %%PAGEBREAK%%
Em seguida: o título "## Gabarito comentado" e, para cada questão, UMA justificativa SUCINTA (1 linha) no formato: "1. B — <justificativa curta> (SOBRENOME, ano, p. X)". Enxuto.`}
function toFileContent(files){const arr=(files||[]).map(f=>{
  // Normalizar media_type para tipos aceitos pela Anthropic
  let mt=f.media_type||'application/pdf';
  if(mt==='application/octet-stream'||mt==='application/msword'){
    if(f.name&&f.name.match(/\.docx?$/i))mt='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  }
  return{type:'document',source:{type:'base64',media_type:mt,data:f.data}};
});if(arr.length)arr[arr.length-1].cache_control={type:'ephemeral'};return arr;}
function txt(res){const b=((res&&res.content)||[]).find(x=>x.type==='text');return b?b.text:''}
// Renderiza markdown com quebra de página onde houver o marcador %%PAGEBREAK%% (gabarito comentado na última página).
function mdPaged(text){const parts=String(text||'').split(/%%PAGEBREAK%%/),out=[];parts.forEach((pt,i)=>{if(i>0)out.push(new Paragraph({children:[new PageBreak()]}));out.push(...mdToDocx(pt.trim()))});return out}
async function gerarDPM(inp,files,tipo){const prompt=tipo==='teorico'?promptDPMTeorico(inp):promptDPMLiterario(inp);const res=await client.messages.create({..._iaExtra(inp,MODEL),max_tokens:12000,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:prompt}]}]});const text=txt(res),si=text.search(/==\s*VERS[ÃA]O\s+PROFESSOR\s*==|(?:^|\n)#{1,3}\s*Quest[ãa]o-Norteadora/i),aR=si>0?text.slice(0,si):text,pR=si>0?text.slice(si):'',aT=aR.replace(/==\s*VERS[ÃA]O\s+ALUNOS\s*==/i,'').trim(),pT=pR.replace(/==\s*VERS[ÃA]O\s+PROFESSOR\s*==/i,'').trim(),isM=(inp.disciplina||'').toLowerCase().includes('metodologia'),docName=tipo==='teorico'?'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Teórico':isM?'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Demonstrativo':'DOCUMENTO DE PARÁGRAFOS MÍNIMOS · DPM Literário',label=tipo==='teorico'?'DPM TEÓRICO':isM?'DPM DEMONSTRATIVO':'DPM LITERÁRIO';const children=[makeHeader(inp.disciplina,inp.semana,docName),p(''),p(`${label} — VERSÃO ALUNOS`,{bold:true,size:28,sb:200}),p(''),...mdToDocx(aT)];if(pT)children.push(...separador(),faixaConf(),p(''),...mdToDocx(pT));return Packer.toBase64String(makeDoc(children))}
async function gerarQuizDoc(inp,files){const res=await client.messages.create({..._iaExtra(inp,MODEL_FAST),max_tokens:2000,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptQuiz(inp)}]}]});return Packer.toBase64String(makeDoc(mdPaged(txt(res))))}
async function gerarBimestralDoc(inp,files){const res=await client.messages.create({..._iaExtra(inp,MODEL_FAST),max_tokens:1400,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:promptBimestral(inp)}]}]});return Packer.toBase64String(makeDoc(mdPaged(txt(res))))}
async function gerarOtimizador(inp){const res=await client.messages.create({..._iaExtra(inp,MODEL),max_tokens:1500,system:SYS_B,messages:[{role:'user',content:`${inp.origem==='painel'?'Material do Painel CLA.\n':''}Queixa/objetivo: ${inp.queixa||'não especificado'}\n\n${inp.prompt}\n\nEntregar: 1) DIAGNÓSTICO 2) VERSÃO OTIMIZADA 3) O QUE MUDOU E POR QUÊ`}]});return txt(res)}
async function gerarDevolutiva(inp,files){const papel=inp.papel==='banca'?'banca (avaliativa)':'orientador (formativa)',fase={inicio:'início',andamento:'andamento',concluido:'concluído (pré-banca)'}[inp.fase]||'',nivel={artigo:'artigo/TCC',dissertacao:'dissertação',tese:'tese'}[inp.nivel]||inp.nivel;const res=await client.messages.create({..._iaExtra(inp,MODEL),max_tokens:2500,system:SYS_C,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:`Papel: ${papel}${inp.papel!=='banca'?' · Fase: '+fase:''} · Nível: ${nivel}\n${inp.foco?'Foco: '+inp.foco:''}\n${inp.contexto?'Contexto/trechos:\n'+inp.contexto:''}\n${files&&files.length?'Trabalho anexado acima.':'Usar contexto/trechos fornecidos.'}\n\nCritérios (por peso): 1) Cumprimento dos objetivos 2) Originalidade 3) Fundamentação teórica 4) Correção conceitual · Clareza · Consistência · ABNT\n\nEstrutura:\n1 — Leitura geral (2–4 frases)\n2 — Pontos por critério (só onde há algo a dizer)\n3 — Apontamentos cirúrgicos: trecho → problema → sugestão → fonte\n4 — Prioridades (2–3 providências)\n5 — Próximo passo`}]}]});return txt(res)}
async function gerarDevolutivaDoc(inp,files){const text=await gerarDevolutiva(inp,files),papel=inp.papel==='banca'?'Banca':'Orientador',fase={inicio:'Início',andamento:'Andamento',concluido:'Concluído'}[inp.fase]||'',nivel={artigo:'Artigo/TCC',dissertacao:'Dissertação',tese:'Tese'}[inp.nivel]||inp.nivel;return Packer.toBase64String(makeDoc([p(`Devolutiva · ${papel}${fase?' · '+fase:''} · ${nivel}`,{size:18,color:'555555',sb:0}),p(''),...mdToDocx(text)]))}
// Extração automática de referência ABNT a partir da ficha catalográfica / página de rosto do PDF.
const PROMPT_REF=`Você recebeu um documento acadêmico (PDF ou DOCX). Identifique os dados bibliográficos para montar a referência ABNT.
Priorize, nesta ordem: a FICHA CATALOGRÁFICA (página de créditos / CIP), a página de rosto e o cabeçalho/rodapé do artigo.
Classifique o tipo em "livro" (livro inteiro), "artigo" (artigo de periódico/revista) ou "capitulo" (capítulo de livro com organizador).
Responda APENAS com um objeto JSON válido, sem nenhum texto fora dele, com EXATAMENTE estas chaves (strings):
{"tipo":"","autores":"","titulo":"","subtitulo":"","revista":"","volume":"","numero":"","ano":"","paginas":"","doi":"","cidade":"","editora":"","edicao":"","organizadores":"","livro":""}
Regras: "autores" e "organizadores" no formato ABNT "SOBRENOME, Nome; SOBRENOME, Nome". "edicao" como "2. ed." quando houver (vazio se 1ª ed.). "paginas" no formato "172-181". Use STRING VAZIA para todo dado que NÃO constar no documento — NUNCA invente autor, ano, editora, cidade, páginas ou título.`;
async function extrairRefABNT(files){const res=await client.messages.create({model:MODEL_FAST,max_tokens:600,system:SYS_A,messages:[{role:'user',content:[...toFileContent(files),{type:'text',text:PROMPT_REF}]}]});const t=txt(res),m=t.match(/\{[\s\S]*\}/);if(!m)throw new Error('Não foi possível ler os dados do PDF.');let o;try{o=JSON.parse(m[0])}catch(_){throw new Error('Resposta inesperada ao ler o PDF.')}const K=['tipo','autores','titulo','subtitulo','revista','volume','numero','ano','paginas','doi','cidade','editora','edicao','organizadores','livro'],ref={};K.forEach(k=>{ref[k]=typeof o[k]==='string'?o[k].trim():''});if(['livro','artigo','capitulo'].indexOf(ref.tipo)<0)ref.tipo='livro';return ref;}
app.use(cookieParser());app.use(express.json({limit:'50mb'}));
app.use((req,res,next)=>{const o=req.headers.origin,a=process.env.SITE_URL||o;if(a){res.setHeader('Access-Control-Allow-Origin',a);res.setHeader('Vary','Origin');}res.setHeader('Access-Control-Allow-Credentials','true');res.setHeader('Access-Control-Allow-Methods','GET, POST, OPTIONS');res.setHeader('Access-Control-Allow-Headers','Content-Type');res.setHeader('X-Content-Type-Options','nosniff');res.setHeader('X-Frame-Options','SAMEORIGIN');res.setHeader('Strict-Transport-Security','max-age=31536000; includeSubDomains');res.setHeader('Content-Security-Policy',"default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; img-src 'self' data:; connect-src 'self'; base-uri 'self'; form-action 'self'; frame-ancestors 'self'");if(req.method==='OPTIONS')return res.sendStatus(200);next()});
app.get('/auth-check',(req,res)=>{const d=verifyToken(req.cookies?.cla_session);if(d)return res.json({ok:true,email:d.email,authUrl:null});const p=new URLSearchParams({client_id:env('GOOGLE_CLIENT_ID'),redirect_uri:env('REDIRECT_URI'),response_type:'code',scope:'openid email profile',prompt:'select_account'});res.json({ok:false,email:null,authUrl:`https://accounts.google.com/o/oauth2/v2/auth?${p.toString()}`})});
app.get('/auth-callback',async(req,res)=>{const c=req.query.code;if(!c)return res.status(400).send('Código ausente.');try{const t=await httpsPost('https://oauth2.googleapis.com/token',{code:c,client_id:env('GOOGLE_CLIENT_ID'),client_secret:env('GOOGLE_CLIENT_SECRET'),redirect_uri:env('REDIRECT_URI'),grant_type:'authorization_code'});if(!t.access_token)return res.status(401).send('Falha na autenticação Google.');const u=await httpsGet('https://www.googleapis.com/oauth2/v2/userinfo',t.access_token);if((u.email||'').toLowerCase()!==ALLOWED_EMAIL)return res.status(403).send(`<html><body style="font-family:sans-serif;text-align:center;padding:60px"><h2>Acesso negado</h2><p>Uso exclusivo do Prof. Felipe Vigneron Azevedo.</p><p style="color:#888">${u.email}</p></body></html>`);res.cookie('cla_session',makeToken(u.email),{httpOnly:true,secure:true,sameSite:'lax',maxAge:86400000*7,path:'/'});res.redirect('/')}catch(e){console.error(e);res.status(500).send('Erro interno.')}});
app.post('/cla-api',requireSession,async(req,res)=>{if(!iaRateOk(req.user&&req.user.email))return res.status(429).json({error:'Muitas gerações em 1 minuto. Aguarde alguns segundos.'});const _json0=res.json.bind(res);res.json=(o)=>_json0(o&&typeof o==='object'&&!Array.isArray(o)?Object.assign({uso:_ultimoUso},o):o);_ultimoUso=null;try{const {mode,task,inputs={},files=[]}=req.body||{},inp=inputs,fmt=(inp.formato||'').toUpperCase();inp.budget=(BUDGET[fmt]||{})[inp.nAulas||'']||{};if(mode==='A'){const sem=inp.semana||'X';if(task==='dpm_teorico')return res.json({docx:await gerarDPM(inp,files,'teorico'),filename:`DPM_Teorico_Sem${sem}.docx`,warnings:[]});if(task==='dpm_literario')return res.json({docx:await gerarDPM(inp,files,'literario'),filename:`DPM_Literario_Sem${sem}.docx`,warnings:[]});if(task==='quiz')return res.json({docx:await gerarQuizDoc(inp,files),filename:`Quiz_Sem${sem}.docx`,warnings:[]});if(task==='bimestral')return res.json({docx:await gerarBimestralDoc(inp,files),filename:`Bimestral_Sem${sem}.docx`,warnings:[]})}else if(mode==='B'){if(!inp.prompt)return res.status(400).json({error:'Campo prompt obrigatório.'});return res.json({text:await gerarOtimizador(inp),model:(_ultimoUso&&_ultimoUso.modeloId)||MODEL,warnings:[]})}else if(mode==='C'){if(task==='chat')return res.json({text:await gerarDevolutiva(inp,files),warnings:[]});if(task==='docx')return res.json({docx:await gerarDevolutivaDoc(inp,files),filename:'Devolutiva.docx',warnings:[]})}else if(mode==='ref'){if(!files||!files.length)return res.status(400).json({error:'Nenhum arquivo para ler.'});return res.json({ref:await extrairRefABNT(files),warnings:[]})}res.status(400).json({error:'Modo/tarefa não reconhecido.'})}catch(e){console.error(e);res.status(500).json({error:e.message||'Erro interno.'})}});
app.get('/health',(req,res)=>res.json({ok:true,t:Date.now()})); // leve, p/ keep-warm / monitor
app.get('/',(req,res)=>res.sendFile(path.join(__dirname,'index.html')));
app.listen(PORT,()=>console.log(`CLA running on ${PORT}`));
