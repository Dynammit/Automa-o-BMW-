"""Rede de Concessionarias - painel visual (abre no navegador, roda so no seu computador).

Como usar: de dois cliques em Iniciar_Rede.bat. O navegador abre sozinho.
Este arquivo precisa ficar na mesma pasta de atualizar_rede.py.
"""
import contextlib, io, json, os, re, secrets, shutil, subprocess, sys, threading, time, traceback, webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

APP = Path(__file__).resolve().parent
os.chdir(APP)
sys.path.insert(0, str(APP))
import pandas as pd
import atualizar_rede as R

ENTRADA, MANUAL, SAIDA = APP / "entrada", APP / "manual", APP / "saida"
for p in (ENTRADA, MANUAL, SAIDA):
    p.mkdir(exist_ok=True)
TOKEN = secrets.token_urlsafe(16)
MANUAIS = {  # arquivo -> (descricao, extensao)
    "gac.json": ("GAC (obrigatório todo mês) - resposta 'list-by-nearby' copiada do F12", ".json"),
    "zeekr.json": ("Zeekr (obrigatório todo mês) - resposta '.../store/list' copiada do F12", ".json"),
    "jetour.json": ("Jetour (só se o download automático falhar)", ".json"),
    "caoa_chery.html": ("Caoa Chery (só se falhar) - página salva com Ctrl+S", ".html"),
    "geely.html": ("Geely (só se falhar) - página /contato salva", ".html"),
    "mg.html": ("MG (só se falhar)", ".html"),
    "volvo.html": ("Volvo (só se falhar; o site costuma bloquear com erro 403) - página salva com Ctrl+S", ".html"),
}
ST = {"busy": False, "log": [], "erro": None, "res": None}
LOCK = threading.Lock()


def log(msg):
    with LOCK:
        for ln in str(msg).splitlines():
            if ln.strip():
                ST["log"].append(ln.rstrip())


class _Tee(io.TextIOBase):
    def __init__(self):
        super().__init__(); self.buf = ""

    def write(self, s):
        self.buf += s
        while "\n" in self.buf:
            ln, self.buf = self.buf.split("\n", 1); log(ln)
        return len(s)

    def flush(self):
        if self.buf.strip():
            log(self.buf)
        self.buf = ""


def fmt_dt(ts):
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M")


def lista_atual():
    arqs = [p for p in ENTRADA.glob("*.xlsx") if not p.name.startswith("~$")]
    return max(arqs, key=lambda p: p.stat().st_mtime) if arqs else None


def limpa(df):
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else (round(v, 1) if isinstance(v, float) else v))
    return df.to_dict("records")


def rodar(marcas):
    ST.update(busy=True, erro=None, log=[], res=None)
    try:
        plan = lista_atual()
        if not plan:
            raise RuntimeError("Envie primeiro a planilha da equipe de rede (passo 1).")
        pasta = SAIDA / f"{datetime.now():%Y-%m-%d_%H%M}"
        with contextlib.redirect_stdout(_Tee()):
            r = R.executar(str(plan), marcas or None, None, str(pasta))
        rel = lambda x: str(Path(x).resolve().relative_to(SAIDA.resolve())).replace("\\", "/")
        ST["res"] = {"arquivo": rel(r["arquivo"]), "lista_comparada": rel(r["lista_comparada"]) if r.get("lista_comparada") else None,
                     "falhas": r.get("falhas", []), "sem_destaque": r.get("pintar_nao", []),
                     "abas": {k: limpa(v) for k, v in r["abas"].items()}, "quando": fmt_dt(time.time()), "lista": plan.name}
    except SystemExit as e:
        ST["erro"] = str(e)
    except Exception as e:
        ST["erro"] = f"{e.__class__.__name__}: {e}"
        log(traceback.format_exc())
    finally:
        ST["busy"] = False


def diagnostico():
    ST.update(busy=True, erro=None, log=[])
    try:
        with contextlib.redirect_stdout(_Tee()):
            R.diagnostico_conexao()
    except Exception as e:
        ST["erro"] = f"{e.__class__.__name__}: {e}"
    finally:
        ST["busy"] = False


def start(fn, *a):
    with LOCK:
        if ST["busy"]:
            return False
        ST["busy"] = True
    threading.Thread(target=fn, args=a, daemon=True).start()
    return True


def estado():
    pl = lista_atual()
    man = {n: ({"mod": fmt_dt((MANUAL / n).stat().st_mtime)} if (MANUAL / n).exists() else None) for n in MANUAIS}
    antigas = sorted([p for p in SAIDA.glob("*/*.xlsx") if not p.name.startswith("~$")], key=lambda p: -p.stat().st_mtime)[:10]
    return {"lista": {"nome": pl.name, "mod": fmt_dt(pl.stat().st_mtime)} if pl else None,
            "manuais": [{"nome": n, "desc": MANUAIS[n][0], "arq": man[n], "aviso": next((m["situacao"] for m in R.estado_manuais() if m["arquivo"] == n and m["precisa_atencao"]), "")} for n in MANUAIS],
            "marcas": [{"k": k, "nome": m, "esp": R.ESPERADO.get(k)} for k, m, _, _ in R.MARCAS],
            "busy": ST["busy"], "erro": ST["erro"], "res": ST["res"],
            "antigas": [{"nome": str(p.relative_to(SAIDA)).replace("\\", "/"), "mod": fmt_dt(p.stat().st_mtime)} for p in antigas]}


def abrir_pasta(p):
    try:
        if os.name == "nt":
            os.startfile(str(p))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception:
        pass


def nome_seguro(n):
    return re.sub(r"[\\/:*?\"<>|]", "_", os.path.basename(n)).strip(". ") or "arquivo"


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _ok(self):
        return self.headers.get("X-Token") == TOKEN or parse_qs(urlparse(self.path).query).get("t", [""])[0] == TOKEN

    def _json(self, obj, code=200):
        b = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code); self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/":
            b = PAGE.replace("__TOKEN__", TOKEN).encode("utf-8")
            self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)
            return
        if not self._ok():
            return self._json({"erro": "não autorizado"}, 403)
        q = parse_qs(u.query)
        if u.path == "/api/estado":
            i = int(q.get("log", ["0"])[0]); e = estado()
            with LOCK:
                e["log"] = ST["log"][i:]; e["log_n"] = len(ST["log"])
            return self._json(e)
        if u.path == "/api/baixar":
            p = (SAIDA / q.get("arq", [""])[0]).resolve()
            if SAIDA.resolve() not in p.parents or not p.exists():
                return self._json({"erro": "arquivo não encontrado"}, 404)
            b = p.read_bytes()
            self.send_response(200); self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(p.name)}")
            self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)
            return
        self._json({"erro": "?"}, 404)

    def do_POST(self):
        if not self._ok():
            return self._json({"erro": "não autorizado"}, 403)
        u = urlparse(self.path); q = parse_qs(u.query)
        n = int(self.headers.get("Content-Length") or 0)
        if u.path == "/api/upload":
            tipo, nome = q.get("tipo", [""])[0], nome_seguro(q.get("nome", ["x"])[0])
            if tipo == "lista":
                if not nome.lower().endswith(".xlsx"):
                    return self._json({"erro": "A planilha da rede precisa ser .xlsx"}, 400)
                dest = ENTRADA / nome
            elif tipo in MANUAIS:
                dest = MANUAL / tipo
            else:
                return self._json({"erro": "tipo inválido"}, 400)
            tmp = dest.with_suffix(dest.suffix + ".parte"); left = n
            with open(tmp, "wb") as f:
                while left > 0:
                    c = self.rfile.read(min(1 << 20, left))
                    if not c:
                        break
                    f.write(c); left -= len(c)
            os.replace(tmp, dest)
            return self._json({"ok": True, "nome": dest.name})
        body = json.loads(self.rfile.read(n) or b"{}") if n else {}
        if u.path == "/api/rodar":
            return self._json({"ok": start(rodar, body.get("marcas") or None)})
        if u.path == "/api/diagnostico":
            return self._json({"ok": start(diagnostico)})
        if u.path == "/api/pasta":
            abrir_pasta(SAIDA); return self._json({"ok": True})
        if u.path == "/api/remover_manual":
            if body.get("nome") in MANUAIS:
                (MANUAL / body["nome"]).unlink(missing_ok=True)
            return self._json({"ok": True})
        if u.path == "/api/encerrar":
            self._json({"ok": True})
            threading.Thread(target=lambda: (time.sleep(0.4), os._exit(0)), daemon=True).start()
            return
        self._json({"erro": "?"}, 404)


PAGE = r"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Rede de Concessionárias</title>
<style>
:root{--bg:#f3f5f8;--card:#fff;--tx:#14213a;--mut:#5b6679;--ln:#dfe4ec;--bl:#1c69d4;--bl2:#0653b6;--ok:#0b7a43;--okbg:#e4f5ea;--wa:#8a5a00;--wabg:#fff1d6;--er:#a3201c;--erbg:#fde5e3}
@media(prefers-color-scheme:dark){:root{--bg:#0f1624;--card:#18223a;--tx:#e8edf6;--mut:#9aa7bf;--ln:#2a3857;--bl:#4c93ff;--bl2:#2f78ee;--ok:#5fd292;--okbg:#12362a;--wa:#f1c36b;--wabg:#3a2f12;--er:#ff9b94;--erbg:#40201f}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.45 system-ui,Segoe UI,Arial,sans-serif}
header{background:linear-gradient(90deg,#0c2d6b,#1c69d4);color:#fff;padding:18px 24px;display:flex;justify-content:space-between;align-items:center}
header h1{margin:0;font-size:20px;font-weight:650}header small{opacity:.85}
main{max-width:1150px;margin:0 auto;padding:20px 16px 60px}
.card{background:var(--card);border:1px solid var(--ln);border-radius:14px;padding:18px 20px;margin-bottom:18px}
.card h2{margin:0 0 4px;font-size:17px;display:flex;align-items:center;gap:10px}
.n{display:inline-grid;place-items:center;width:28px;height:28px;border-radius:50%;background:var(--bl);color:#fff;font-weight:700;font-size:14px}
.sub{color:var(--mut);margin:0 0 14px}
.badge{margin-left:auto;font-size:12px;font-weight:650;padding:3px 10px;border-radius:99px}
.b-ok{background:var(--okbg);color:var(--ok)}.b-wa{background:var(--wabg);color:var(--wa)}.b-er{background:var(--erbg);color:var(--er)}.b-no{background:var(--ln);color:var(--mut)}
.drop{border:2px dashed var(--ln);border-radius:12px;padding:16px;text-align:center;color:var(--mut);cursor:pointer}
.drop.over{border-color:var(--bl);background:rgba(28,105,212,.07);color:var(--bl)}.drop b{color:var(--tx);display:block}
button{font:inherit;border:0;border-radius:10px;padding:10px 18px;cursor:pointer;font-weight:600}
.pri{background:var(--bl);color:#fff}.pri:hover{background:var(--bl2)}.sec{background:var(--ln);color:var(--tx)}button:disabled{opacity:.5;cursor:not-allowed}
table{border-collapse:collapse;width:100%;font-size:13.5px}th,td{padding:6px 9px;border-bottom:1px solid var(--ln);text-align:left;vertical-align:top}
th{color:var(--mut);font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.03em;position:sticky;top:0;background:var(--card)}
td.r,th.r{text-align:right}
.alert{padding:9px 12px;border-radius:9px;margin:6px 0;font-size:14px}.alert.aviso{background:var(--wabg);color:var(--wa)}.alert.erro{background:var(--erbg);color:var(--er)}.alert.ok{background:var(--okbg);color:var(--ok)}
.log{background:#0d1526;color:#cfe0ff;border-radius:10px;padding:10px 12px;font:12.5px/1.5 ui-monospace,Consolas,monospace;max-height:200px;overflow:auto;white-space:pre-wrap;margin-top:12px}
.bar{height:6px;background:var(--ln);border-radius:9px;overflow:hidden;margin-top:12px}.bar i{display:block;height:100%;width:40%;background:var(--bl);animation:mv 1.1s infinite ease-in-out}
@keyframes mv{0%{margin-left:-40%}100%{margin-left:100%}}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}.hide{display:none}
.marcas{display:flex;flex-wrap:wrap;gap:8px 16px;margin:6px 0 14px}.marcas label{display:flex;gap:6px;align-items:center;cursor:pointer}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin:14px 0 8px}.tabs button{padding:6px 12px;font-size:13px}.tabs .on{background:var(--bl);color:#fff}
.sc{overflow:auto;max-height:420px;border:1px solid var(--ln);border-radius:10px}
.man td{padding:5px 8px}.man button{padding:2px 9px;font-size:12px}
details summary{cursor:pointer;color:var(--bl);font-weight:600;margin:8px 0}
input[type=text]{font:inherit;padding:7px 10px;border-radius:8px;border:1px solid var(--ln);background:var(--card);color:var(--tx);width:260px}
a.dl{display:inline-block;background:var(--bl);color:#fff;border-radius:10px;padding:10px 18px;font-weight:600;text-decoration:none}
</style></head><body>
<header><div><h1>Rede de Concessionárias</h1><small>Marcas chinesas · comparação com a lista da equipe de rede</small></div><button class="sec" onclick="encerrar()">Encerrar o painel</button></header>
<main>
<div class="card"><h2><span class="n">1</span>Planilha da rede <span class="badge b-no" id="b1">…</span></h2>
<p class="sub">Arraste a planilha mais recente da equipe de rede (a que tem a aba “Lista Operações”). Vale sempre a última enviada.</p>
<div class="drop" id="dz-lista"><b>📥 Planilha da rede (.xlsx)</b><span id="slista">Arraste aqui ou clique para escolher</span></div></div>

<div class="card"><h2><span class="n">2</span>Atualizar as lojas <span class="badge b-no" id="b2">não feito</span></h2>
<p class="sub">Baixa as lojas dos sites oficiais e compara com a lista. Demora alguns minutos (a GWM é a mais lenta). Se uma marca falhar, as outras continuam.</p>
<div class="marcas" id="marcas"></div>
<div class="row"><button class="pri" id="brun" onclick="rodar()">▶ Atualizar agora</button><button class="sec" id="bdiag" onclick="diag()">🔌 Testar conexão</button><button class="sec hide" id="bfalhas" onclick="refazer()">↻ Refazer só as marcas com problema</button><button class="sec" onclick="pasta()">Abrir pasta dos resultados</button></div>
<div class="bar hide" id="bar"><i></i></div><div class="log hide" id="log"></div>
<details style="margin-top:14px"><summary>Arquivos manuais (plano B – quando o site bloqueia o download automático)</summary>
<p class="sub">A <b>GAC</b> precisa deste arquivo todo mês. As outras só se aparecer erro. Arraste o arquivo na linha correspondente.</p>
<table class="man"><tbody id="manuais"></tbody></table></details></div>

<div class="card hide" id="cres"><h2><span class="n">3</span>Resultado <span class="badge b-no" id="b3">—</span></h2>
<p class="sub" id="quando"></p>
<div id="alertas"></div>
<div class="sc" style="max-height:none"><table id="tresumo"></table></div>
<div class="tabs" id="tabs"></div>
<input type="text" id="filtro" placeholder="Buscar nas linhas…" oninput="filtra()" style="margin-bottom:8px">
<div class="sc"><table id="tdet"></table></div>
<div class="row" style="margin-top:14px"><a class="dl" id="dl2" href="#">⬇ Planilha para a equipe de rede (com cores)</a><a class="dl" id="dl" href="#" style="background:var(--ln);color:var(--tx)">Relatório de conferência (só para você)</a></div>
<p class="sub" style="margin-top:12px"><b>Planilha com cores:</b> verde = possível loja nova (acrescentada no fim da lista), vermelho = operação que não aparece no site; a coluna STATUS (PROGRAMA) explica cada linha. Marcas com erro, leitura incompleta ou dados antigos ficam sem destaque. Tudo aqui é “possível”: confirme com a equipe de rede antes de alterar a lista. O grupo de cada loja é uma estimativa (pelo nome, e-mail ou CNPJ).</p></div>

<div class="card"><h2>Resultados anteriores</h2><table><tbody id="antigas"></tbody></table></div>
</main>
<script>
const TK="__TOKEN__";let logN=0,poll=null,E=null,aba="Para confirmar";
const $=id=>document.getElementById(id);
async function api(p,b){const o={headers:{"X-Token":TK}};if(b!==undefined){o.method="POST";o.body=JSON.stringify(b)}return (await fetch(p,o)).json()}
function badge(id,c,t){const e=$(id);e.className="badge "+c;e.textContent=t}
function zone(id,tipo){const z=$(id),i=document.createElement("input");i.type="file";i.style.display="none";z.after(i);
 z.onclick=()=>i.click();i.onchange=()=>up(tipo,i.files[0]);z.ondragover=e=>{e.preventDefault();z.classList.add("over")};
 z.ondragleave=()=>z.classList.remove("over");z.ondrop=e=>{e.preventDefault();z.classList.remove("over");up(tipo,e.dataTransfer.files[0])}}
async function up(tipo,f){if(!f)return;const r=await fetch("/api/upload?tipo="+tipo+"&nome="+encodeURIComponent(f.name),{method:"POST",headers:{"X-Token":TK},body:f});const j=await r.json();if(j.erro)alert(j.erro);carregar()}
zone("dz-lista","lista");
let montou=false;
async function carregar(){E=await api("/api/estado?log="+logN);logN=E.log_n;render()}
function render(){
 $("slista").textContent=E.lista?`✔ ${E.lista.nome} (enviada em ${E.lista.mod})`:"Arraste aqui ou clique para escolher";
 badge("b1",E.lista?"b-ok":"b-wa",E.lista?"ok":"falta enviar");
 if(!montou){$("marcas").innerHTML=E.marcas.map(m=>`<label><input type="checkbox" value="${m.k}" checked>${m.nome}</label>`).join("");montou=true}
 const mb=$("manuais");if(!mb.dataset.ok){mb.dataset.ok=1;mb.innerHTML=E.manuais.map(m=>`<tr><td><b>${m.nome}</b><br><small style="color:var(--mut)">${m.desc}</small>${m.aviso?`<br><small style="color:var(--wa);font-weight:600">⚠ ${m.aviso}: atualize antes de rodar</small>`:""}</td><td id="ms-${m.nome}"></td><td><div class="drop" id="dz-${m.nome}" style="padding:6px 10px">arraste aqui</div></td></tr>`).join("");E.manuais.forEach(m=>zone("dz-"+m.nome,m.nome))}
 E.manuais.forEach(m=>{$("ms-"+m.nome).innerHTML=m.arq?`✔ ${m.arq.mod} <button class="sec" onclick="remMan('${m.nome}')">remover</button>`:"—"});
 $("brun").disabled=E.busy||!E.lista;$("bdiag").disabled=E.busy;$("bar").classList.toggle("hide",!E.busy);
 if(E.log&&E.log.length){const l=$("log");l.classList.remove("hide");E.log.forEach(t=>{const d=document.createElement("div");d.textContent=t;if(/ERRO/.test(t))d.style.color="#ff9b94";l.appendChild(d)});l.scrollTop=l.scrollHeight}
 $("antigas").innerHTML=E.antigas.length?E.antigas.map(a=>`<tr><td>${a.nome}</td><td>${a.mod}</td><td><a href="/api/baixar?arq=${encodeURIComponent(a.nome)}&t=${TK}">baixar</a></td></tr>`).join(""):`<tr><td style="color:var(--mut)">Nenhum ainda.</td></tr>`;
 if(E.busy){badge("b2","b-wa","em andamento…");return}
 if(E.erro){$("cres").classList.remove("hide");$("alertas").innerHTML=`<div class="alert erro">⛔ ${E.erro}</div>`;badge("b2","b-er","erro");return}
 if(E.res)resultado()}
function esc(v){return String(v??"").replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]))}
function resultado(){const r=E.res,A=r.abas,res=A["Resumo geral"],av=A["Avisos"];
 $("cres").classList.remove("hide");$("quando").textContent=`Gerado em ${r.quando} com a lista ${r.lista}.`;
 const er=av.filter(a=>a.Tipo=="ERRO"),at=av.filter(a=>a.Tipo=="ATENÇÃO");
 badge("b2",er.length?"b-wa":"b-ok",er.length?er.length+" marca(s) com erro":"concluído");badge("b3",er.length?"b-wa":"b-ok",er.length?"revisar erros":"ok");
 $("alertas").innerHTML=er.map(a=>`<div class="alert erro">⛔ <b>${esc(a.Marca)}</b>: ${esc(a.Texto)}</div>`).join("")+at.map(a=>`<div class="alert aviso">⚠️ <b>${esc(a.Marca)}</b>: ${esc(a.Texto)}</div>`).join("")||`<div class="alert ok">✔ Todas as marcas rodaram sem erro.</div>`;
 const c=["Marca","Lojas no site","Variação vs. mês anterior","Operações na lista","Iguais","Endereço divergente","Possíveis novas","Possíveis fechadas","Situação"];
 $("tresumo").innerHTML="<tr>"+c.map((x,i)=>`<th class="${i&&i<8?"r":""}">${x}</th>`).join("")+"</tr>"+res.map(l=>"<tr>"+c.map((x,i)=>`<td class="${i&&i<8?"r":""}"${x=="Situação"&&String(l[x]).startsWith("ERRO")?` style="color:var(--er)"`:(x=="Situação"&&String(l[x]).startsWith("ATENÇÃO")?` style="color:var(--wa)"`:"")}>${esc(l[x])}</td>`).join("")+"</tr>").join("");
 const nomes=["Para confirmar","Possiveis lojas novas","Possiveis fechadas","Endereco divergente","Outros grupos","Qualidade da lista","Avisos"];
 $("tabs").innerHTML=nomes.map(n=>`<button class="sec ${n==aba?"on":""}" onclick="aba='${n}';resultado()">${n.replace("Possiveis","Possíveis").replace("Endereco","Endereço")} (${A[n].length})</button>`).join("");
 const d=A[aba];if(!d.length){$("tdet").innerHTML="<tr><td style='color:var(--mut)'>Nada nesta aba.</td></tr>"}else{const k=Object.keys(d[0]);
  $("tdet").innerHTML="<tr>"+k.map(x=>`<th>${esc(x)}</th>`).join("")+"</tr>"+d.map(l=>"<tr>"+k.map(x=>`<td>${esc(l[x])}</td>`).join("")+"</tr>").join("")}
 $("dl").href="/api/baixar?arq="+encodeURIComponent(r.arquivo)+"&t="+TK;
 const d2=$("dl2");if(r.lista_comparada){d2.style.display="";d2.href="/api/baixar?arq="+encodeURIComponent(r.lista_comparada)+"&t="+TK}else{d2.style.display="none"}
 $("bfalhas").classList.toggle("hide",!(r.falhas&&r.falhas.length));filtra()}
function filtra(){const f=$("filtro").value.toLowerCase();document.querySelectorAll("#tdet tr").forEach((r,i)=>{if(i)r.style.display=r.textContent.toLowerCase().includes(f)?"":"none"})}
async function rodar(){const ms=[...document.querySelectorAll("#marcas input:checked")].map(i=>i.value);if(!ms.length)return alert("Marque pelo menos uma marca.");
 logN=0;$("log").innerHTML="";aba="Para confirmar";await api("/api/rodar",{marcas:ms.length==E.marcas.length?[]:ms});espera()}
function espera(){clearInterval(poll);poll=setInterval(async()=>{await carregar();if(!E.busy)clearInterval(poll)},1200);carregar()}
async function diag(){logN=0;$("log").innerHTML="";await api("/api/diagnostico",{});espera()}
function refazer(){const f=(E.res&&E.res.falhas)||[];document.querySelectorAll("#marcas input").forEach(i=>i.checked=f.includes(i.value));rodar()}
async function remMan(n){await api("/api/remover_manual",{nome:n});carregar()}
async function pasta(){await api("/api/pasta",{})}
async function encerrar(){if(confirm("Encerrar o painel?")){await api("/api/encerrar",{});document.body.innerHTML="<p style='padding:40px;font:18px system-ui'>Painel encerrado. Pode fechar esta aba.</p>"}}
carregar();
</script></body></html>"""


def main():
    srv = None
    for p in range(8780, 8810):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", p), H); port = p; break
        except OSError:
            continue
    if not srv:
        sys.exit("Não consegui abrir uma porta local.")
    url = f"http://127.0.0.1:{port}/"
    print(f"Painel aberto em {url}\nNão feche esta janela enquanto estiver usando o painel.")
    if "--sem-navegador" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()