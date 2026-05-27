#!/usr/bin/env python3
"""
Contact2Sale — App Unificado v1
Preparador de Planilha + Importador de Vendedores

Rodar local : python3 c2s_app.py
Deploy      : renomeie para main.py + requirements.txt vazio

Credenciais (edite aqui):
"""
import http.server, json, urllib.request, urllib.error
import threading, webbrowser, sys, time, os, secrets

APP_USER = "admin"
APP_PASS  = "contato2024"   # ← mude aqui

PORT = int(os.environ.get("PORT", 8742))
SESSIONS: dict = {}

def check_session(cookie_header: str) -> bool:
    if not cookie_header: return False
    for part in cookie_header.split(";"):
        k, _, v = part.strip().partition("=")
        if k.strip() == "sid" and v.strip() in SESSIONS: return True
    return False

def new_session() -> str:
    sid = secrets.token_hex(32)
    SESSIONS[sid] = True
    return sid

# ─────────────────────────────────────────────────────────────────────
LOGIN_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>Contact2Sale — Login</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#f5f4f0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  display:flex;align-items:center;justify-content:center;min-height:100vh}
.box{background:#fff;border:1px solid #e2e0d8;border-radius:14px;padding:2.5rem 2rem;
  width:360px;box-shadow:0 4px 24px rgba(0,0,0,.07)}
.logo{width:44px;height:44px;background:#2563eb;border-radius:11px;display:flex;
  align-items:center;justify-content:center;margin:0 auto 1.25rem}
.logo svg{width:22px;height:22px;fill:none;stroke:#fff;stroke-width:2}
h1{font-size:20px;font-weight:700;letter-spacing:-.03em;text-align:center;color:#1a1917}
p{font-size:13px;color:#6b6860;text-align:center;margin-top:4px;margin-bottom:1.5rem}
label{font-size:12px;font-weight:600;color:#6b6860;text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:4px}
input{width:100%;height:42px;border:1px solid #ccc9be;border-radius:8px;padding:0 12px;
  font-size:14px;outline:none;margin-bottom:12px;transition:border-color .15s,box-shadow .15s}
input:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.12)}
button{width:100%;height:42px;background:#2563eb;border:none;border-radius:8px;color:#fff;
  font-size:14px;font-weight:600;cursor:pointer;margin-top:4px;transition:background .15s}
button:hover{background:#1d4ed8}
.err{background:#fef2f2;border:1px solid #fecaca;color:#dc2626;border-radius:8px;
  padding:8px 12px;font-size:13px;margin-bottom:12px;display:none}
</style></head>
<body>
<div class="box">
  <div class="logo"><svg viewBox="0 0 24 24"><path d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z"/></svg></div>
  <h1>Contact2Sale</h1>
  <p>Preparador & Importador de Vendedores</p>
  <div class="err" id="err">Usuário ou senha incorretos.</div>
  <label>Usuário</label>
  <input type="text" id="u" autocomplete="username" placeholder="usuario"/>
  <label>Senha</label>
  <input type="password" id="p" autocomplete="current-password" placeholder="••••••••"/>
  <button onclick="login()">Entrar</button>
</div>
<script>
async function login(){
  const u=document.getElementById('u').value.trim();
  const p=document.getElementById('p').value;
  const r=await fetch('/auth',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({u,p})});
  if(r.ok){location.href='/';}else{document.getElementById('err').style.display='block';}
}
document.addEventListener('keydown',e=>{if(e.key==='Enter')login();});
</script></body></html>"""

# ─────────────────────────────────────────────────────────────────────
APP_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Contact2Sale</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f5f4f0;--surface:#fff;--s2:#f0efe9;
  --border:#e2e0d8;--border2:#ccc9be;
  --text:#1a1917;--text2:#6b6860;--text3:#9c9a93;
  --blue:#2563eb;--blue-bg:#eff4ff;--blue-b:#bfcffe;
  --green:#16a34a;--green-bg:#f0fdf4;--green-b:#bbf7d0;
  --red:#dc2626;--red-bg:#fef2f2;--red-b:#fecaca;
  --amber:#d97706;--amber-bg:#fffbeb;--amber-b:#fde68a;
  --purple:#7c3aed;--purple-bg:#f5f3ff;--purple-b:#ddd6fe;
  --r:10px;--rs:6px;
  --mono:'JetBrains Mono','Fira Code',monospace;
}
body{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;line-height:1.6;min-height:100vh}

/* ── Topbar ── */
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:0 2rem;height:52px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:200}
.tlogo{width:28px;height:28px;background:var(--blue);border-radius:7px;display:flex;align-items:center;justify-content:center}
.tlogo svg{width:16px;height:16px;fill:none;stroke:#fff;stroke-width:2}
.ttitle{font-size:14px;font-weight:600}
.tsep{color:var(--border2);margin:0 4px}
.tbadge{font-size:11px;padding:2px 8px;background:var(--green-bg);color:var(--green);border:1px solid var(--green-b);border-radius:20px;font-weight:500}
.tlogout{margin-left:auto;font-size:12px;padding:4px 12px;border:1px solid var(--border2);border-radius:6px;background:none;cursor:pointer;color:var(--text2)}
.tlogout:hover{background:var(--s2)}

/* ── Nav Tabs ── */
.nav{background:var(--surface);border-bottom:1px solid var(--border);display:flex;padding:0 2rem;gap:0}
.nav-tab{padding:12px 20px;font-size:13px;font-weight:500;color:var(--text2);cursor:pointer;border-bottom:2px solid transparent;transition:all .15s;display:flex;align-items:center;gap:7px}
.nav-tab:hover{color:var(--text);background:var(--s2)}
.nav-tab.active{color:var(--blue);border-bottom-color:var(--blue);background:transparent}
.nav-tab .nav-icon{font-size:15px}

/* ── Pages ── */
.page{display:none;max-width:860px;margin:0 auto;padding:2rem 1.5rem 4rem}
.page.active{display:block}
h1{font-size:22px;font-weight:700;letter-spacing:-.03em}
.sub{font-size:13px;color:var(--text2);margin-top:3px;margin-bottom:0}

/* ── Cards ── */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);overflow:hidden;margin-top:16px}
.card.disabled{opacity:.4;pointer-events:none}
.ch{display:flex;align-items:center;gap:10px;padding:14px 18px;border-bottom:1px solid var(--border);background:var(--s2)}
.sn{width:22px;height:22px;border-radius:50%;border:1.5px solid var(--border2);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;color:var(--text2);flex-shrink:0;transition:all .2s}
.sn.active{background:var(--blue);border-color:var(--blue);color:#fff}
.sn.active.purple{background:var(--purple);border-color:var(--purple)}
.sn.done{background:var(--green);border-color:var(--green);color:#fff}
.ct{font-size:11px;font-weight:600;letter-spacing:.06em;text-transform:uppercase}
.cb{padding:18px}

/* ── Forms ── */
input[type=text],input[type=password]{width:100%;height:40px;border:1px solid var(--border2);border-radius:var(--rs);padding:0 12px;font-size:13px;color:var(--text);background:var(--surface);outline:none;transition:border-color .15s,box-shadow .15s}
input[type=text]:focus,input[type=password]:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.frow{display:flex;gap:8px;align-items:center}
.frow input{flex:1}
label.lbl{font-size:12px;font-weight:600;color:var(--text2);display:block;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}

/* ── Buttons ── */
.btn{height:38px;padding:0 16px;border:1px solid var(--border2);border-radius:var(--rs);background:var(--surface);color:var(--text);font-size:13px;font-weight:500;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:all .15s;white-space:nowrap}
.btn:hover{background:var(--s2)}.btn:active{transform:scale(.98)}.btn:disabled{opacity:.5;cursor:not-allowed;transform:none}
.bp{background:var(--blue);border-color:var(--blue);color:#fff}.bp:hover{background:#1d4ed8}
.bpurple{background:var(--purple);border-color:var(--purple);color:#fff}.bpurple:hover{background:#6d28d9}
.bgreen{background:var(--green-bg);border-color:var(--green-b);color:var(--green)}.bgreen:hover{background:var(--green-b)}
.bsm{height:30px;padding:0 12px;font-size:12px}

/* ── Alerts / Status ── */
.msg{margin-top:10px;font-size:13px;padding:8px 12px;border-radius:var(--rs);display:none}
.msg.ok{background:var(--green-bg);color:var(--green);border:1px solid var(--green-b);display:block}
.msg.err{background:var(--red-bg);color:var(--red);border:1px solid var(--red-b);display:block}
.msg.info{background:var(--blue-bg);color:var(--blue);border:1px solid var(--blue-b);display:block}
.alert{border-radius:var(--rs);padding:10px 14px;font-size:13px;display:flex;gap:8px;align-items:flex-start;line-height:1.5;margin-bottom:12px}
.alert-amber{background:var(--amber-bg);border:1px solid var(--amber-b);color:var(--amber)}
.alert-blue{background:var(--blue-bg);border:1px solid var(--blue-b);color:var(--blue)}
.alert-green{background:var(--green-bg);border:1px solid var(--green-b);color:var(--green)}

/* ── Chips ── */
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
.chip{font-size:12px;padding:3px 10px;background:var(--blue-bg);color:var(--blue);border:1px solid var(--blue-b);border-radius:20px;font-weight:500}
.chip-purple{background:var(--purple-bg);color:var(--purple);border-color:var(--purple-b)}
.chip-rm{display:inline-flex;align-items:center;gap:6px}
.chip-rm button{background:none;border:none;cursor:pointer;color:inherit;font-size:14px;line-height:1;padding:0;opacity:.7}
.chip-rm button:hover{opacity:1;color:var(--red)}

/* ── Drop zone ── */
.dz{border:2px dashed var(--border2);border-radius:var(--rs);padding:2rem 1rem;text-align:center;cursor:pointer;transition:all .15s;color:var(--text3)}
.dz:hover,.dz.over{border-color:var(--blue);background:var(--blue-bg);color:var(--blue)}
.dz svg{width:32px;height:32px;margin-bottom:8px;opacity:.5}
.dz p{font-size:13px;margin-top:4px}
.dz small{font-size:12px;opacity:.7}
.fileok{display:none;align-items:center;gap:10px;padding:10px 14px;background:var(--green-bg);border:1px solid var(--green-b);border-radius:var(--rs);margin-top:10px}
.fileok .fn{font-size:13px;font-weight:600;color:var(--green);flex:1}

/* ── Tabs dentro da página ── */
.itabs{display:flex;gap:0;border:1px solid var(--border);border-radius:var(--rs);overflow:hidden;margin-bottom:14px}
.itab{flex:1;padding:8px 12px;text-align:center;font-size:13px;font-weight:500;color:var(--text2);cursor:pointer;background:var(--s2);border-right:1px solid var(--border);transition:all .15s}
.itab:last-child{border-right:none}
.itab.active{background:var(--purple);color:#fff}
.itab-panel{display:none}.itab-panel.active{display:block}
textarea{width:100%;border:1px solid var(--border2);border-radius:var(--rs);padding:10px 12px;font-size:13px;font-family:var(--mono);color:var(--text);background:var(--surface);outline:none;resize:vertical;min-height:140px;transition:border-color .15s}
textarea:focus{border-color:var(--purple);box-shadow:0 0 0 3px rgba(124,58,237,.1)}

/* ── Grid / Editor ── */
.grid-wrap{overflow-x:auto;border:1px solid var(--border);border-radius:var(--rs)}
table.grid{width:100%;border-collapse:collapse;font-size:13px}
table.grid thead tr{background:var(--s2)}
table.grid th{padding:8px 10px;text-align:left;font-size:11px;font-weight:600;color:var(--text2);border-bottom:1px solid var(--border);border-right:1px solid var(--border);text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
table.grid th:last-child{border-right:none}
table.grid td{padding:0;border-bottom:1px solid var(--border);border-right:1px solid var(--border)}
table.grid td:last-child{border-right:none}
table.grid tr:last-child td{border-bottom:none}
table.grid td input.cell{width:100%;height:34px;border:none;padding:0 8px;font-size:13px;font-family:inherit;color:var(--text);background:transparent;outline:none}
table.grid td input.cell:focus{background:var(--blue-bg);box-shadow:inset 0 0 0 2px var(--blue)}
table.grid td input.cell.missing{background:var(--red-bg)}
table.grid td select.tsel{width:100%;height:34px;border:none;padding:0 6px;font-size:12px;color:var(--text);background:transparent;outline:none;cursor:pointer}
table.grid td select.tsel:focus{background:var(--blue-bg)}
table.grid td select.tsel.is-default{color:var(--text3);font-style:italic}
table.grid td select.tsel.is-team{color:var(--purple);font-weight:500}
table.grid td select.tsel.is-matched{color:var(--green);font-weight:500}
table.grid td.rdel{width:32px;text-align:center}
table.grid td.rdel button{background:none;border:none;color:var(--text3);cursor:pointer;font-size:16px;line-height:1;padding:4px}
table.grid td.rdel button:hover{color:var(--red)}
.result-toolbar{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap}
.result-title{font-size:12px;font-weight:600;color:var(--text2);text-transform:uppercase;letter-spacing:.05em;flex:1}

/* ── Stats ── */
.stats-bar{display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap}
.stat{background:var(--s2);border:1px solid var(--border);border-radius:var(--rs);padding:8px 14px;font-size:12px;color:var(--text2)}
.stat strong{color:var(--text);font-size:15px;display:block;font-weight:700;line-height:1.2}

/* ── Importador: progress / log ── */
.pbar{margin-top:14px;display:none}
.pmeta{display:flex;justify-content:space-between;font-size:12px;color:var(--text2);margin-bottom:6px}
.ptrack{height:6px;background:var(--s2);border-radius:3px;overflow:hidden;border:1px solid var(--border)}
.pfill{height:100%;background:var(--blue);border-radius:3px;transition:width .3s;width:0%}
.sgrid{display:none;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px}
.sc{background:var(--s2);border:1px solid var(--border);border-radius:var(--rs);padding:12px 14px}
.sc .v{font-size:26px;font-weight:700;letter-spacing:-.03em;color:var(--text);line-height:1}
.sc .l{font-size:11px;color:var(--text2);margin-top:4px;text-transform:uppercase;letter-spacing:.05em;font-weight:500}
.sc.ok .v{color:var(--green)}.sc.er .v{color:var(--red)}.sc.wa .v{color:var(--amber)}
.logsec{margin-top:14px;display:none}
.loghdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
.logttl{font-size:11px;font-weight:600;color:var(--text2);text-transform:uppercase;letter-spacing:.06em}
.logbox{background:#0f1117;border-radius:var(--rs);padding:12px 14px;max-height:320px;overflow-y:auto;font-family:var(--mono);font-size:12px;line-height:1.7}
.ll{display:flex;gap:10px;align-items:flex-start}
.lt{color:#4b5563;flex-shrink:0;font-size:11px;margin-top:1px}
.lok{color:#4ade80}.ler{color:#f87171}.lwa{color:#fbbf24}.lin{color:#60a5fa}.ldm{color:#6b7280}
.spin{animation:sp 1s linear infinite;display:inline-block}
@keyframes sp{from{transform:rotate(0)}to{transform:rotate(360deg)}}

/* ── Username box ── */
.ubox{background:var(--s2);border:1px solid var(--border);border-radius:var(--rs);padding:12px 14px;margin-bottom:14px;font-size:12px;color:var(--text2);line-height:1.9}
.ubox strong{color:var(--text)}
.ubox code{font-family:var(--mono);background:var(--border);padding:1px 5px;border-radius:3px;color:var(--blue);font-size:11px}
.cobox{margin-top:12px;background:var(--s2);border:1px solid var(--border);border-radius:var(--rs);padding:12px 14px;display:none}
.coname{font-size:15px;font-weight:600}
.cometa{font-size:11px;font-family:var(--mono);color:var(--text3);margin-top:2px}
</style>
</head>
<body>

<div class="topbar">
  <div class="tlogo"><svg viewBox="0 0 24 24"><path d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z"/></svg></div>
  <span class="ttitle">Contact2Sale</span>
  <span class="tbadge">🟢 Online</span>
  <button class="tlogout" onclick="logout()">Sair</button>
</div>

<div class="nav">
  <div class="nav-tab active" id="nav-prep" onclick="switchPage('prep')">
    <span class="nav-icon">✨</span> Preparador de Planilha
  </div>
  <div class="nav-tab" id="nav-imp" onclick="switchPage('imp')">
    <span class="nav-icon">🚀</span> Importador de Vendedores
  </div>
</div>

<!-- ══════════════════════════════════════════════════════════════════
     PÁGINA 1 — PREPARADOR
═══════════════════════════════════════════════════════════════════ -->
<div class="page active" id="page-prep">
  <h1>Preparador de Planilha</h1>
  <p class="sub">Recebe dados em qualquer formato e gera o arquivo pronto para importar.</p>

  <!-- P1: Empresa -->
  <div class="card" id="p1c1">
    <div class="ch"><div class="sn active purple" id="p1b1">1</div><span class="ct">Empresa e equipes</span></div>
    <div class="cb">
      <div class="alert alert-blue">ℹ️ Informe a empresa do cliente e as equipes (se houver). Sem equipe definida, o padrão será o nome da empresa.</div>
      <label class="lbl">Nome da empresa (padrão)</label>
      <input type="text" id="p1-empresa" placeholder="Ex: Onboarding Matheus"/>
      <div style="margin-top:14px">
        <label class="lbl">Equipes disponíveis <span style="font-weight:400;text-transform:none">(opcional)</span></label>
        <div class="frow" style="margin-top:4px">
          <input type="text" id="p1-eq" placeholder="Ex: Equipe Carlos, Equipe Ramon"/>
          <button class="btn bpurple bsm" onclick="p1AddTeam()">+ Adicionar</button>
        </div>
        <div class="chips" id="p1-teams" style="margin-top:10px">
          <span style="font-size:12px;color:var(--text3)">Nenhuma equipe — empresa será usada como padrão.</span>
        </div>
      </div>
      <div style="margin-top:16px;display:flex;justify-content:flex-end">
        <button class="btn bpurple" onclick="p1Confirm()">Continuar →</button>
      </div>
    </div>
  </div>

  <!-- P1: Dados -->
  <div class="card disabled" id="p1c2">
    <div class="ch"><div class="sn" id="p1b2">2</div><span class="ct">Dados dos usuários</span></div>
    <div class="cb">
      <div class="itabs">
        <div class="itab active" onclick="p1SetTab('arq')">📎 Arquivo (Excel / TXT)</div>
        <div class="itab" onclick="p1SetTab('txt')">✏️ Colar texto / WhatsApp</div>
      </div>
      <div class="itab-panel active" id="p1-tab-arq">
        <div class="dz" id="p1-dz" onclick="document.getElementById('p1-fi').click()"
          ondragover="event.preventDefault();this.classList.add('over')"
          ondragleave="this.classList.remove('over')" ondrop="p1Drop(event)">
          <svg fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
            <path d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/>
          </svg>
          <p>Clique ou arraste .xlsx, .xls, .csv ou .txt</p>
          <small>Qualquer formatação — detecta campos automaticamente</small>
        </div>
        <input type="file" id="p1-fi" accept=".xlsx,.xls,.csv,.txt" style="display:none" onchange="p1HandleFile(this.files[0])"/>
      </div>
      <div class="itab-panel" id="p1-tab-txt">
        <textarea id="p1-txt" placeholder="Cole aqui os dados — WhatsApp, e-mail, qualquer formato:

João Silva, joao@empresa.com, (11) 99999-1111
Maria Souza | maria@empresa.com | 11988887777
Nome: Carlos / Email: carlos@emp.com / Tel: 11977776666"></textarea>
        <div style="font-size:12px;color:var(--text2);margin-top:6px;padding:6px 10px;background:var(--s2);border-radius:var(--rs)">
          💡 O sistema identifica nome, e-mail e telefone automaticamente. Revise no editor abaixo.
        </div>
        <div style="margin-top:10px;display:flex;justify-content:flex-end">
          <button class="btn bpurple" onclick="p1ProcessText()">✨ Processar texto</button>
        </div>
      </div>
    </div>
  </div>

  <!-- P1: Editor -->
  <div class="card disabled" id="p1c3">
    <div class="ch"><div class="sn" id="p1b3">3</div><span class="ct">Revisar e exportar</span></div>
    <div class="cb" id="p1-result">
      <div id="p1-alert"></div>
      <div class="stats-bar" id="p1-stats"></div>
      <div class="result-toolbar">
        <span class="result-title">✏️ Editor — ajuste antes de exportar</span>
        <button class="btn bsm" onclick="p1AddRow()">+ Linha</button>
        <button class="btn bsm bp" onclick="p1Export()">⬇ Exportar .xlsx</button>
        <button class="btn bsm bpurple" onclick="p1SendToImporter()">🚀 Enviar para Importador</button>
      </div>
      <div class="grid-wrap"><table class="grid"><thead id="p1-gh"></thead><tbody id="p1-gb"></tbody></table></div>
      <button class="btn bsm" style="margin-top:8px" onclick="p1AddRow()">+ Adicionar linha</button>
    </div>
  </div>
</div>

<!-- ══════════════════════════════════════════════════════════════════
     PÁGINA 2 — IMPORTADOR
═══════════════════════════════════════════════════════════════════ -->
<div class="page" id="page-imp">
  <h1>Importador de Vendedores</h1>
  <p class="sub">Token → planilha → editor → importação com log completo.</p>

  <!-- I1: Token -->
  <div class="card" id="ic1">
    <div class="ch"><div class="sn active" id="ib1">1</div><span class="ct">Token de autenticação</span></div>
    <div class="cb">
      <div class="frow">
        <input type="password" id="i-tok" placeholder="Cole seu Bearer token..." autocomplete="off"/>
        <button class="btn bp" id="i-btnval" onclick="iValidate()">✔ Validar</button>
      </div>
      <div class="msg" id="i-tmsg"></div>
      <div class="cobox" id="i-cobox">
        <div class="coname" id="i-coname"></div>
        <div class="cometa" id="i-cometa"></div>
        <div id="i-sublabel" style="font-size:12px;font-weight:600;color:var(--text2);margin:10px 0 6px;text-transform:uppercase;letter-spacing:.04em;display:none">Equipes disponíveis:</div>
        <div class="chips" id="i-chips"></div>
      </div>
    </div>
  </div>

  <!-- I2: Planilha -->
  <div class="card disabled" id="ic2">
    <div class="ch"><div class="sn" id="ib2">2</div><span class="ct">Planilha de usuários</span></div>
    <div class="cb">
      <div class="alert alert-amber">⚠ Carregue qualquer planilha e edite antes de importar. Colunas obrigatórias: <strong>name</strong>, <strong>email</strong>.</div>
      <div class="dz" id="i-dz" onclick="document.getElementById('i-fi').click()"
        ondragover="event.preventDefault();this.classList.add('over')"
        ondragleave="this.classList.remove('over')" ondrop="iDrop(event)">
        <svg fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
          <path d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/>
        </svg>
        <p>Clique ou arraste .xlsx / .csv</p>
      </div>
      <input type="file" id="i-fi" accept=".xlsx,.xls,.csv" style="display:none" onchange="iHandleFile(this.files[0])"/>
      <div class="fileok" id="i-fok">
        ✅ <span class="fn" id="i-fn"></span>
        <span style="font-size:12px;color:var(--green)" id="i-fc"></span>
        <button class="btn bsm" onclick="document.getElementById('i-fi').click()">Trocar</button>
      </div>
      <div id="i-editor-wrap" style="display:none;margin-top:16px">
        <div class="result-toolbar">
          <span class="result-title">✏️ Editor</span>
          <button class="btn bsm" onclick="iAddRow()">+ Linha</button>
          <button class="btn bsm bgreen" onclick="iConfirm()">✔ Confirmar e continuar</button>
        </div>
        <div id="i-hint" style="font-size:12px;padding:8px 12px;border-radius:var(--rs);margin-bottom:10px;border:1px solid transparent"></div>
        <div class="grid-wrap"><table class="grid"><thead id="i-gh"></thead><tbody id="i-gb"></tbody></table></div>
        <button class="btn bsm" style="margin-top:8px" onclick="iAddRow()">+ Adicionar linha</button>
      </div>
    </div>
  </div>

  <!-- I3: Importar -->
  <div class="card disabled" id="ic3">
    <div class="ch"><div class="sn" id="ib3">3</div><span class="ct">Importar vendedores</span></div>
    <div class="cb">
      <div class="ubox">
        <strong>Lógica de username:</strong><br>
        1ª → <code id="i-ex1">matheusoliveira.onboarding</code> (primeiro+último nome . primeira palavra da empresa)<br>
        2ª → <code>email do contato</code><br>
        3ª → <code id="i-ex3">matheusoliveirakx.onboarding</code> (+ 2 letras aleatórias)
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
        <button class="btn bp" id="i-btnimp" onclick="iStartImport()">▶ Iniciar importação</button>
        <button class="btn bsm" id="i-btndl" onclick="iDlLog()" style="display:none">⬇ Baixar log</button>
        <button class="btn bsm" id="i-btncl" onclick="iClLog()" style="display:none">🗑 Limpar</button>
      </div>
      <div class="pbar" id="i-pbar">
        <div class="pmeta"><span id="i-plbl">Aguardando...</span><span id="i-ppct">0%</span></div>
        <div class="ptrack"><div class="pfill" id="i-pfill"></div></div>
      </div>
      <div class="sgrid" id="i-sgrid">
        <div class="sc">    <div class="v" id="i-stot">0</div><div class="l">Total</div></div>
        <div class="sc ok"><div class="v" id="i-sok">0</div> <div class="l">Sucesso</div></div>
        <div class="sc er"><div class="v" id="i-ser">0</div> <div class="l">Falha</div></div>
        <div class="sc wa"><div class="v" id="i-sre">0</div> <div class="l">Retentativas</div></div>
      </div>
      <div class="logsec" id="i-logsec">
        <div class="loghdr"><span class="logttl">Log de execução</span></div>
        <div class="logbox" id="i-logbox"></div>
      </div>
    </div>
  </div>
</div>

<script>
// ══════════════════════════════════════════════════════════════════
// SHARED
// ══════════════════════════════════════════════════════════════════
const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const ts  = () => new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/[^a-z0-9]/g,'');
const isEmail = s => /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(String(s||'').trim());
const isPhone = s => /^[\\d\\s\\(\\)\\-\\+]{7,}$/.test(String(s||'').trim());
const firstWord = name => norm((name||'').trim().split(/\\s+/)[0] || name || 'empresa');

function switchPage(name){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  document.getElementById('nav-'+name).classList.add('active');
}

function sDone(id){const e=document.getElementById(id);e.classList.remove('active');e.classList.add('done');e.textContent='✓';}
function sActive(id){document.getElementById(id).classList.add('active');}
function setMsg(id,html,type){const e=document.getElementById(id);e.className='msg '+type;e.innerHTML=html;}
async function logout(){await fetch('/logout',{method:'POST'});location.href='/login';}

async function proxy(method,url,body){
  const r=await fetch('/proxy',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({method,url,token:iToken,body:body||null})});
  return r.json();
}

// ══════════════════════════════════════════════════════════════════
// PREPARADOR
// ══════════════════════════════════════════════════════════════════
let p1Empresa='', p1Teams=[], p1Rows=[];

function p1AddTeam(){
  const inp=document.getElementById('p1-eq');
  const val=inp.value.trim(); if(!val)return;
  val.split(',').map(v=>v.trim()).filter(Boolean).forEach(v=>{
    if(!p1Teams.find(t=>norm(t)===norm(v))) p1Teams.push(v);
  });
  inp.value=''; p1RenderTeams();
}
function p1RemoveTeam(i){p1Teams.splice(i,1);p1RenderTeams();}
function p1RenderTeams(){
  const el=document.getElementById('p1-teams');
  if(!p1Teams.length){el.innerHTML='<span style="font-size:12px;color:var(--text3)">Nenhuma equipe — empresa será usada como padrão.</span>';return;}
  el.innerHTML=p1Teams.map((t,i)=>`<span class="chip chip-purple chip-rm">${esc(t)}<button onclick="p1RemoveTeam(${i})">×</button></span>`).join('');
}
function p1Confirm(){
  p1Empresa=document.getElementById('p1-empresa').value.trim();
  if(!p1Empresa){alert('Informe o nome da empresa.');return;}
  sDone('p1b1');
  document.getElementById('p1c2').classList.remove('disabled');
  sActive('p1b2');
  document.getElementById('p1c2').scrollIntoView({behavior:'smooth',block:'start'});
}
function p1SetTab(name){
  document.querySelectorAll('.itab').forEach((t,i)=>t.classList.toggle('active',['arq','txt'][i]===name));
  document.querySelectorAll('.itab-panel').forEach((p,i)=>p.classList.toggle('active',['arq','txt'][i]===name));
}
function p1Drop(e){e.preventDefault();document.getElementById('p1-dz').classList.remove('over');if(e.dataTransfer.files[0])p1HandleFile(e.dataTransfer.files[0]);}
function p1HandleFile(file){
  if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  const reader=new FileReader();
  if(ext==='txt'){reader.onload=e=>p1ParseRaw(e.target.result);reader.readAsText(file,'UTF-8');}
  else{reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      p1ParseSheet(XLSX.utils.sheet_to_json(ws,{defval:'',header:1}));
    }catch(err){alert('Erro: '+err.message);}
  };reader.readAsArrayBuffer(file);}
}
function p1ParseSheet(raw){
  if(!raw.length)return;
  const first=raw[0].map(c=>String(c||'').toLowerCase().trim());
  const hasH=first.some(c=>['name','nome','email','telefone','phone'].includes(c));
  let data=hasH?raw.slice(1):raw;
  data=data.filter(r=>r.some(c=>String(c||'').trim()));
  let iN=-1,iE=-1,iP=-1,iT=-1;
  if(hasH){first.forEach((h,i)=>{
    if(['name','nome'].includes(h))iN=i;
    else if(h==='email')iE=i;
    else if(['phone','phone1','telefone','celular','fone'].includes(h))iP=i;
    else if(['team','equipe','time','gerente'].includes(h))iT=i;
  });}
  if(iN<0||iE<0){
    const sample=data.slice(0,5);const cols=(raw[0]||[]).length;
    for(let c=0;c<cols;c++){
      const vals=sample.map(r=>String(r[c]||'').trim()).filter(Boolean);
      if(vals.filter(isEmail).length>=vals.length*.4&&iE<0)iE=c;
      else if(vals.filter(isPhone).length>=vals.length*.4&&iP<0)iP=c;
    }
    for(let c=0;c<cols;c++){if(c!==iE&&c!==iP&&c!==iT&&iN<0)iN=c;}
  }
  p1Rows=data.map(r=>p1BuildRow(
    iN>=0?String(r[iN]||'').trim():'',
    iE>=0?String(r[iE]||'').trim():'',
    iP>=0?String(r[iP]||'').trim():'',
    iT>=0?String(r[iT]||'').trim():''
  )).filter(r=>r.name||r.email);
  if(!p1Rows.length){alert('Não encontrei dados válidos.');return;}
  p1ShowResult();
}
function p1ProcessText(){
  const txt=document.getElementById('p1-txt').value.trim();
  if(!txt){alert('Cole algum texto primeiro.');return;}
  p1ParseRaw(txt);
}
function p1ParseRaw(text){
  const lines=text.split(/\\n/).map(l=>l.trim()).filter(Boolean);
  const sepRx=/[,|;/\\t]/;
  const structured=lines.filter(l=>sepRx.test(l));
  const contacts=[];
  if(structured.length>=lines.length*.5){
    lines.forEach(line=>{
      const parts=line.split(sepRx).map(p=>p.trim()).filter(Boolean);
      let name='',email='',phone='',team='';
      parts.forEach(p=>{
        const clean=p.replace(/^(nome|name|email|e-mail|tel|telefone|fone|celular|equipe|team)\\s*[:=]\\s*/i,'').trim();
        if(isEmail(clean))email=clean;
        else if(isPhone(clean))phone=clean;
        else if(!name)name=clean;
        else if(!team)team=clean;
      });
      if(name||email)contacts.push({name,email,phone,team});
    });
  }else{
    let cur={name:'',email:'',phone:'',team:''};
    lines.forEach(line=>{
      const clean=line.replace(/^(nome|name|email|e-mail|tel|telefone|fone|celular|equipe|team)\\s*[:=]\\s*/i,'').trim();
      if(isEmail(clean))cur.email=clean;
      else if(isPhone(clean))cur.phone=clean;
      else if(p1Teams.find(t=>norm(clean).includes(norm(t))))cur.team=clean;
      else{if(cur.name||cur.email){contacts.push({...cur});cur={name:'',email:'',phone:'',team:''};}cur.name=clean;}
    });
    if(cur.name||cur.email)contacts.push({...cur});
  }
  if(!contacts.length){alert('Não consegui identificar contatos.');return;}
  p1Rows=contacts.map(c=>p1BuildRow(c.name,c.email,c.phone,c.team));
  p1ShowResult();
}
function p1BuildRow(name,email,phone,teamHint){
  let team='';
  if(teamHint&&p1Teams.length){
    const m=p1Teams.find(t=>norm(teamHint).includes(norm(t))||norm(t).includes(norm(teamHint)));
    team=m||'';
  }
  if(!team)team=p1Empresa;
  return{name,email,phone1:phone,team};
}
function p1ShowResult(){
  document.getElementById('p1c3').classList.remove('disabled');
  sDone('p1b2');sActive('p1b3');
  p1BuildGrid();p1BuildStats();
  document.getElementById('p1c3').scrollIntoView({behavior:'smooth',block:'start'});
}
function p1BuildStats(){
  const we=p1Rows.filter(r=>r.email).length,wp=p1Rows.filter(r=>r.phone1).length;
  const wt=p1Rows.filter(r=>r.team&&r.team!==p1Empresa).length,ne=p1Rows.filter(r=>!r.email).length;
  document.getElementById('p1-stats').innerHTML=
    `<div class="stat"><strong>${p1Rows.length}</strong>contatos</div>`+
    `<div class="stat"><strong>${we}</strong>com e-mail</div>`+
    `<div class="stat"><strong>${wp}</strong>com telefone</div>`+
    `<div class="stat"><strong>${wt}</strong>com equipe</div>`+
    (ne?`<div class="stat" style="color:var(--red)"><strong>${ne}</strong>sem e-mail ⚠</div>`:'');
  document.getElementById('p1-alert').innerHTML=ne
    ?`<div class="alert alert-amber">⚠ ${ne} contato(s) sem e-mail — preencha manualmente abaixo.</div>`
    :`<div class="alert alert-green">✔ Todos os contatos têm e-mail. Revise e exporte.</div>`;
}
function p1BuildGrid(){
  document.getElementById('p1-gh').innerHTML='<tr><th>Nome</th><th>E-mail</th><th>Telefone</th><th>Equipe</th><th style="width:32px"></th></tr>';
  p1RenderBody();
}
function p1RenderBody(){
  const tb=document.getElementById('p1-gb');tb.innerHTML='';
  p1Rows.forEach((r,i)=>{
    let opts=`<option value="${esc(p1Empresa)}" ${r.team===p1Empresa?'selected':''}>${esc(p1Empresa)} (padrão)</option>`;
    p1Teams.forEach(t=>{if(t!==p1Empresa)opts+=`<option value="${esc(t)}" ${r.team===t?'selected':''}>${esc(t)}</option>`;});
    const tr=document.createElement('tr');
    tr.innerHTML=
      `<td><input class="cell" value="${esc(r.name)}" onchange="p1Rows[${i}].name=this.value"/></td>`+
      `<td><input class="cell${!r.email?' missing':''}" value="${esc(r.email)}" onchange="p1Rows[${i}].email=this.value" placeholder="${!r.email?'⚠ obrigatório':''}"/></td>`+
      `<td><input class="cell" value="${esc(r.phone1)}" onchange="p1Rows[${i}].phone1=this.value"/></td>`+
      `<td><select class="tsel ${r.team===p1Empresa?'is-default':'is-team'}" onchange="p1SetTeam(${i},this)">${opts}</select></td>`+
      `<td class="rdel"><button onclick="p1DelRow(${i})">×</button></td>`;
    tb.appendChild(tr);
  });
}
function p1SetTeam(i,sel){p1Rows[i].team=sel.value;sel.className='tsel '+(sel.value===p1Empresa?'is-default':'is-team');}
function p1DelRow(i){p1Rows.splice(i,1);p1RenderBody();p1BuildStats();}
function p1AddRow(){
  p1Rows.push({name:'',email:'',phone1:'',team:p1Empresa});
  p1RenderBody();
  document.querySelectorAll('#p1-gb tr:last-child input.cell')[0]?.focus();
}
function p1Export(){
  const valid=p1Rows.filter(r=>r.name||r.email);
  if(!valid.length){alert('Nenhum dado para exportar.');return;}
  const data=[['name','email','phone1','team'],...valid.map(r=>[r.name,r.email,r.phone1,r.team])];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws['!cols']=[{wch:30},{wch:35},{wch:18},{wch:30}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Usuarios');
  XLSX.writeFile(wb,'c2s_'+norm(p1Empresa)+'_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
function p1SendToImporter(){
  // Exporta e muda de aba
  p1Export();
  setTimeout(()=>{
    switchPage('imp');
    alert('Planilha exportada! Agora carregue o arquivo na aba Importador.');
  },300);
}

// ══════════════════════════════════════════════════════════════════
// IMPORTADOR
// ══════════════════════════════════════════════════════════════════
let iToken='', iSlug='', iSubs=[], iRows=[], iLogLines=[];

function iLog(msg,type='in'){
  const b=document.getElementById('i-logbox'),t=ts();
  iLogLines.push('['+t+'] '+msg);
  const d=document.createElement('div');d.className='ll';
  const cls={ok:'lok',er:'ler',wa:'lwa',in:'lin',dm:'ldm'}[type]||'lin';
  d.innerHTML='<span class="lt">'+t+'</span><span class="'+cls+'">'+esc(msg)+'</span>';
  b.appendChild(d);b.scrollTop=b.scrollHeight;
}
function fuzzyMatch(val,subs){
  const n=norm(val);if(!n)return null;
  return subs.find(s=>norm(s.company_name).includes(n))||null;
}
async function iValidate(){
  const raw=document.getElementById('i-tok').value.trim();
  if(!raw)return; iToken=raw;
  const btn=document.getElementById('i-btnval');
  btn.disabled=true;btn.innerHTML='<span class="spin">↻</span> Validando...';
  setMsg('i-tmsg','⏳ Conectando...','info');
  const res=await proxy('GET','https://api.contact2sale.com/integration/me',null);
  if(!res.ok){
    setMsg('i-tmsg','✖ Falha: '+esc(res.error||JSON.stringify(res.body||{})),'err');
    btn.disabled=false;btn.innerHTML='✔ Validar';return;
  }
  const d=res.body||{};
  iSubs=d.sub_companies||[];
  iSlug=firstWord(d.company_name||'empresa');
  setMsg('i-tmsg','✔ Token válido — '+esc(d.company_name||''),'ok');
  document.getElementById('i-coname').textContent=d.company_name||'';
  document.getElementById('i-cometa').textContent='ID: '+(d.company_id||'')+'  ·  slug: '+iSlug;
  if(iSubs.length){
    document.getElementById('i-sublabel').style.display='block';
    document.getElementById('i-chips').innerHTML=iSubs.map(s=>'<span class="chip">'+esc(s.company_name)+'</span>').join('');
  }
  document.getElementById('i-cobox').style.display='block';
  document.getElementById('i-ex1').textContent='matheusoliveira.'+iSlug;
  document.getElementById('i-ex3').textContent='matheusoliveirakx.'+iSlug;
  sDone('ib1');
  document.getElementById('ic2').classList.remove('disabled');
  sActive('ib2');
}
function iDrop(e){e.preventDefault();document.getElementById('i-dz').classList.remove('over');if(e.dataTransfer.files[0])iHandleFile(e.dataTransfer.files[0]);}
function iHandleFile(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=evt=>{
    try{
      const wb=XLSX.read(evt.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{defval:'',header:1});
      if(!raw.length)throw new Error('Planilha vazia.');
      const headers=raw[0].map(h=>String(h||'').trim());
      const data=raw.slice(1).filter(r=>r.some(c=>String(c||'').trim()));
      const idx=names=>names.map(n=>headers.findIndex(h=>norm(h)===norm(n))).find(i=>i>=0)??-1;
      const iN=idx(['name','nome']),iE=idx(['email']),iP=idx(['phone1','phone','telefone','celular']),iT=idx(['team','equipe','time']);
      iRows=data.map(r=>({
        name:  iN>=0?String(r[iN]||'').trim():'',
        email: iE>=0?String(r[iE]||'').trim():'',
        phone1:iP>=0?String(r[iP]||'').trim():'',
        team:  iT>=0?String(r[iT]||'').trim():'',
      })).map(r=>{
        const m=r.team?fuzzyMatch(r.team,iSubs):null;
        return{...r,_cid:m?m.company_id:(iSubs[0]?.company_id||''),_cname:m?m.company_name:(iSubs[0]?.company_name||''),_matched:!!m};
      });
      if(!iRows.length)throw new Error('Nenhum dado válido.');
      document.getElementById('i-fn').textContent=file.name;
      document.getElementById('i-fc').textContent=iRows.length+' linhas';
      document.getElementById('i-fok').style.display='flex';
      document.getElementById('i-dz').style.display='none';
      iBuildEditor();
      document.getElementById('i-editor-wrap').style.display='block';
    }catch(e){alert('Erro: '+e.message);}
  };
  reader.readAsArrayBuffer(file);
}
function iBuildEditor(){
  const un=iRows.filter(r=>r.team&&!r._matched).length;
  const h=document.getElementById('i-hint');
  if(un>0){h.textContent='⚠ '+un+' linha(s) com equipe não reconhecida.';h.style.cssText='background:var(--amber-bg);border-color:var(--amber-b);color:var(--amber)';}
  else{h.textContent='✔ Equipes mapeadas. Edite se necessário.';h.style.cssText='background:var(--green-bg);border-color:var(--green-b);color:var(--green)';}
  document.getElementById('i-gh').innerHTML='<tr><th>Nome</th><th>E-mail</th><th>Telefone</th><th>Equipe</th><th style="width:32px"></th></tr>';
  iRenderBody();
}
function iRenderBody(){
  const tb=document.getElementById('i-gb');tb.innerHTML='';
  iRows.forEach((r,i)=>{
    let opts='<option value="">— sem equipe —</option>';
    iSubs.forEach(s=>{opts+=`<option value="${esc(s.company_id)}" ${s.company_id===r._cid?'selected':''}>${esc(s.company_name)}</option>`;});
    const cls=r._cid?(r._matched?'is-matched':'is-team'):'is-default';
    const tr=document.createElement('tr');
    tr.innerHTML=
      `<td><input class="cell" value="${esc(r.name)}" onchange="iRows[${i}].name=this.value"/></td>`+
      `<td><input class="cell" value="${esc(r.email)}" onchange="iRows[${i}].email=this.value"/></td>`+
      `<td><input class="cell" value="${esc(r.phone1)}" onchange="iRows[${i}].phone1=this.value"/></td>`+
      `<td><select class="tsel ${cls}" onchange="iSetTeam(${i},this)">${opts}</select></td>`+
      `<td class="rdel"><button onclick="iDelRow(${i})">×</button></td>`;
    tb.appendChild(tr);
  });
}
function iSetTeam(i,sel){
  const s=iSubs.find(x=>x.company_id===sel.value);
  iRows[i]._cid=s?s.company_id:'';iRows[i]._cname=s?s.company_name:'';iRows[i]._matched=true;
  sel.className='tsel '+(s?'is-matched':'is-default');
}
function iDelRow(i){iRows.splice(i,1);iRenderBody();}
function iAddRow(){
  iRows.push({name:'',email:'',phone1:'',team:'',_cid:iSubs[0]?.company_id||'',_cname:iSubs[0]?.company_name||'',_matched:false});
  iRenderBody();
  document.querySelectorAll('#i-gb tr:last-child input.cell')[0]?.focus();
}
function iConfirm(){
  iRows=iRows.filter(r=>r.name&&r.email);
  if(!iRows.length){alert('Nenhuma linha com name + email.');return;}
  iRenderBody();
  const h=document.getElementById('i-hint');
  h.textContent='✔ '+iRows.length+' usuário(s) prontos.';
  h.style.cssText='background:var(--green-bg);border-color:var(--green-b);color:var(--green)';
  sDone('ib2');
  document.getElementById('ic3').classList.remove('disabled');
  sActive('ib3');
  document.getElementById('ic3').scrollIntoView({behavior:'smooth',block:'start'});
}
function genUser(name,email,slug,att){
  const p=name.toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/[^a-z\\s]/g,'').trim().split(/\\s+/);
  const base=(p[0]||'user')+(p[p.length-1]||'x');
  if(att===1)return base+'.'+slug;
  if(att===2)return email;
  const rnd=Array.from({length:2},()=>'abcdefghijklmnopqrstuvwxyz'[Math.floor(Math.random()*26)]).join('');
  return base+rnd+'.'+slug;
}
async function iStartImport(){
  const valid=iRows.filter(r=>r.name&&r.email);
  if(!valid.length||!iToken){alert('Confirme os dados primeiro.');return;}
  const btn=document.getElementById('i-btnimp');
  btn.disabled=true;btn.innerHTML='<span class="spin">↻</span> Importando...';
  document.getElementById('i-pbar').style.display='block';
  document.getElementById('i-sgrid').style.display='grid';
  document.getElementById('i-logsec').style.display='block';
  document.getElementById('i-btndl').style.display='';
  document.getElementById('i-btncl').style.display='';
  let ok=0,err=0,retries=0;
  const total=valid.length;
  iLog('▶ Iniciando — '+total+' usuário(s)','in');
  for(let i=0;i<valid.length;i++){
    const u=valid[i];
    const slug=u._cname?firstWord(u._cname):iSlug;
    const cid=u._cid||iSubs[0]?.company_id||'';
    const pct=Math.round(((i+1)/total)*100);
    document.getElementById('i-pfill').style.width=pct+'%';
    document.getElementById('i-ppct').textContent=pct+'%';
    document.getElementById('i-plbl').textContent=(i+1)+'/'+total+': '+u.name;
    if(u._cname)iLog('  → '+u._cname,'dm');
    for(let att=1;att<=3;att++){
      const uname=genUser(u.name,u.email,slug,att);
      if(!uname){iLog('✖ '+u.name+' — sem username','er');err++;break;}
      if(att>1){retries++;iLog('  ↻ Tentativa '+att+'/3 — '+uname,'wa');}
      const body={name:u.name,email:u.email,username:uname,company_id:cid};
      if(u.phone1)body.phone1=u.phone1;
      const res=await proxy('POST','https://api.contact2sale.com/integration/sellers',body);
      if(res.ok||res.status===200||res.status===201){iLog('✔ '+u.name+' → '+uname,'ok');ok++;break;}
      const em=res.body?.message||res.body?.error||res.error||'HTTP '+res.status;
      const conflict=res.status===409||res.status===422||['username','exist','already','taken'].some(w=>String(em).toLowerCase().includes(w));
      if(conflict&&att<3){iLog('  ⚠ "'+uname+'" indisponível — próxima...','wa');}
      else{iLog('✖ '+u.name+' — '+em,'er');err++;break;}
    }
    document.getElementById('i-stot').textContent=i+1;
    document.getElementById('i-sok').textContent=ok;
    document.getElementById('i-ser').textContent=err;
    document.getElementById('i-sre').textContent=retries;
    await new Promise(r=>setTimeout(r,150));
  }
  const color=err===0?'#16a34a':ok===0?'#dc2626':'#d97706';
  document.getElementById('i-pfill').style.background=color;
  document.getElementById('i-plbl').textContent='Concluído — '+ok+' sucesso · '+err+' falha';
  iLog('━━ Finalizado: '+ok+' sucesso · '+err+' falha · '+retries+' retentativas ━━','dm');
  btn.disabled=false;btn.innerHTML='✔ Concluído';
}
function iDlLog(){
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([iLogLines.join('\\n')],{type:'text/plain'}));
  a.download='c2s_import_'+new Date().toISOString().slice(0,10)+'.log';a.click();
}
function iClLog(){document.getElementById('i-logbox').innerHTML='';iLogLines=[];}

document.getElementById('i-tok').addEventListener('keydown',e=>{if(e.key==='Enter')iValidate();});
document.getElementById('p1-empresa').addEventListener('keydown',e=>{if(e.key==='Enter')p1Confirm();});
document.getElementById('p1-eq').addEventListener('keydown',e=>{if(e.key==='Enter')p1AddTeam();});
</script>
</body></html>"""


# ─────────────────────────────────────────────────────────────────────
class H(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _send(self, code, ctype, body_bytes):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", len(body_bytes))
        self.end_headers()
        self.wfile.write(body_bytes)

    def do_GET(self):
        cookies = self.headers.get("Cookie", "")
        if self.path in ("/", "/index.html"):
            if not check_session(cookies):
                self.send_response(302); self.send_header("Location", "/login"); self.end_headers(); return
            self._send(200, "text/html;charset=utf-8", APP_HTML.encode())
        elif self.path == "/login":
            self._send(200, "text/html;charset=utf-8", LOGIN_HTML.encode())
        else:
            self._send(404, "text/plain", b"Not found")

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw    = self.rfile.read(length)
        cookies = self.headers.get("Cookie", "")

        if self.path == "/auth":
            try:
                data = json.loads(raw)
                if secrets.compare_digest(data.get("u",""), APP_USER) and secrets.compare_digest(data.get("p",""), APP_PASS):
                    sid = new_session()
                    self.send_response(200)
                    self.send_header("Set-Cookie", f"sid={sid}; HttpOnly; SameSite=Strict; Path=/")
                    self.send_header("Content-Type", "application/json")
                    self.end_headers(); self.wfile.write(b'{"ok":true}')
                else:
                    self._send(401, "application/json", b'{"ok":false}')
            except Exception as ex:
                self._send(400, "application/json", json.dumps({"error":str(ex)}).encode())
            return

        if self.path == "/logout":
            for part in cookies.split(";"):
                k,_,v = part.strip().partition("=")
                if k.strip()=="sid": SESSIONS.pop(v.strip(),None)
            self.send_response(200); self.send_header("Set-Cookie","sid=; Max-Age=0; Path=/"); self.end_headers()
            return

        if self.path == "/proxy":
            if not check_session(cookies):
                self._send(401, "application/json", b'{"error":"unauthorized"}'); return
            try:
                q=json.loads(raw); method=q.get("method","GET").upper(); url=q.get("url","")
                tok=q.get("token",""); payload=q.get("body",None)
                hdrs={"Authorization":f"Bearer {tok}","Content-Type":"application/json","Accept":"application/json"}
                bdata=json.dumps(payload).encode() if payload is not None else None
                req=urllib.request.Request(url,data=bdata,headers=hdrs,method=method)
                try:
                    with urllib.request.urlopen(req,timeout=30) as resp:
                        result={"ok":True,"status":resp.status,"body":json.loads(resp.read().decode())}
                except urllib.error.HTTPError as e:
                    st=e.code
                    try: rb=json.loads(e.read().decode())
                    except: rb={}
                    result={"ok":False,"status":st,"body":rb,"error":rb.get("message") or rb.get("error") or f"HTTP {st}"}
            except Exception as ex:
                result={"ok":False,"status":0,"error":str(ex),"body":{}}
            out=json.dumps(result).encode()
            self._send(200,"application/json",out)
            return

        self._send(404,"text/plain",b"Not found")


def main():
    srv = http.server.HTTPServer(("0.0.0.0", PORT), H)
    print(f"""
  ╔══════════════════════════════════════════════╗
  ║   Contact2Sale — App Unificado v1            ║
  ╠══════════════════════════════════════════════╣
  ║  ✨ Preparador de Planilha                   ║
  ║  🚀 Importador de Vendedores                 ║
  ╠══════════════════════════════════════════════╣
  ║  Usuário : {APP_USER:<30} ║
  ║  Senha   : {APP_PASS:<30} ║
  ╚══════════════════════════════════════════════╝

  ✅  http://127.0.0.1:{PORT}
  Ctrl+C para encerrar
""")
    threading.Thread(
        target=lambda: (time.sleep(1), webbrowser.open(f"http://127.0.0.1:{PORT}")),
        daemon=True
    ).start()
    try: srv.serve_forever()
    except KeyboardInterrupt: print("\n  Encerrado."); sys.exit(0)

if __name__ == "__main__":
    main()
