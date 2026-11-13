#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ATUALIZAR REDE - baixa as lojas de TODAS as marcas chinesas e compara com a lista da equipe de rede.

COMO USAR (uma vez por mês), no terminal do VS Code, dentro da pasta onde estão este arquivo e a planilha da rede:

    python atualizar_rede.py

O que ele faz, sozinho:
  1) procura na pasta a planilha da rede (a que tem a aba "Lista Operações");
  2) baixa as lojas de cada marca nos sites oficiais (e guarda cada lista em CSV na pasta saida_AAAA-MM-DD);
  3) compara cada marca com a "Lista Operações";
  4) grava UMA planilha só:  saida_AAAA-MM-DD/Comparacao_Rede_AAAA-MM-DD.xlsx

Abas da planilha:
  Resumo geral              uma linha por marca (quantas batem, quantas possíveis novas, possíveis fechadas...)
  Possiveis lojas novas     lojas que estão no site, são de um grupo que já está na lista, e não estão na lista
  Possiveis fechadas        operações da lista que não aparecem mais no site
  Endereco divergente       mesma cidade e grupo, mas endereço/coordenada diferente (conferir)
  Outros grupos             lojas do site de grupos que não estão no relatório
  Avisos                    erros e observações de cada marca

Se uma marca falhar (site mudou, sem internet...), as outras continuam e o aviso aparece na planilha.

Opções (todas opcionais):
  python atualizar_rede.py --lista MinhaPlanilha.xlsx      usa essa planilha em vez de procurar
  python atualizar_rede.py --marcas BYD,GWM                roda só essas marcas
  python atualizar_rede.py --sem-download saida_2026-11-06 não baixa nada; compara usando os CSVs de uma pasta anterior

Arquivos manuais (só quando um site bloqueia o download automático) - coloque na pasta "manual":
  manual/gac.json          GAC: no site da GAC, F12 > Rede > list-by-nearby > Resposta > copiar; cole num arquivo gac.json
  manual/jetour.json       Jetour: abra https://jetourbr.com/wp-admin/admin-ajax.php?action=jetour_lojas_near e salve como jetour.json
  manual/caoa_chery.html, geely.html, mg.html, volvo.html
                           página salva com Ctrl+S ("Página da Web, somente HTML"); para Geely salve a página /contato
O script usa esses arquivos quando o download direto falha (na GAC, usa o arquivo se ele existir).

ONDE EDITAR QUANDO PRECISAR: só a seção "CONFIGURAÇÕES" logo abaixo (grupos, apelidos, distâncias).
"""
import sys, os, re, io, csv, json, math, time, glob, html, uuid, datetime, unicodedata, argparse, traceback
import urllib.request, urllib.error, urllib.parse, ssl, socket, random
try:                                  # usa o repositório de certificados do Windows (resolve o erro de certificado em redes corporativas)
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass
import numpy as np
import pandas as pd

# =====================================================================================================
# CONFIGURAÇÕES  (edite aqui; o resto do arquivo normalmente não precisa mudar)
# =====================================================================================================
ABA_LISTA = "Lista Operações"
COLUNAS_LISTA = ["NOME DO GRUPO", "MARCA", "CIDADE", "UF", "ENDEREÇO", "LATITUDE", "LONGITUDE"]

# Distância (km) para considerar que uma loja do site é a mesma operação da lista
KM_PADRAO = 3.0
KM_BYD = 5.0            # BYD: a lista tem coordenadas menos precisas
KM_MESMA_LOJA_BYD = 0.2  # BYD: venda e pós-venda a menos de 200 m = mesma loja física

# BYD: pedaço do e-mail da loja -> grupo (acrescente linhas se aparecer um grupo novo)
DOMINIOS_BYD = {"grupoiesa": "Iesa", "gruposaga": "Saga", "bydbarigui": "Barigui", "grupobarigui": "Barigui",
                "ravierabyd": "Raviera", "gruporaviera": "Raviera", "grandbrasil": "Grand Brasil",
                "mclartymaia": "McLarty Maia", "caltabiano": "McLarty Maia", "jsl.com": "Simpar",
                "originalautos": "Simpar", "original-veiculos": "Simpar"}

# Marcas em que só entram lojas com ponto próprio (ignora lojas dentro de shopping/outlet, tanto no site quanto na lista)
MARCAS_SEM_SHOPPING = ["GWM"]
PADRAO_SHOPPING = r"shopping|outlet|\bmall\b|\bpiso\b|\bloja\s+[a-z]\d"   # palavra no nome/endereço, ou loja tipo "L26"/"F313"

# Quantidade aproximada de lojas que cada site mostra (só para avisar quando vier algo muito diferente)
ESPERADO = {"BYD": 292, "GWM": 125, "OMODA_JAECOO": 110, "JETOUR": 60, "GAC": 76, "CAOA_CHERY": 160, "LEAPMOTOR": 92,
            "DENZA": 21, "VOLVO": 51, "GEELY": 57, "MG": 40, "ZEEKR": 14}

# Apelidos: nome comercial que aparece no site -> nome do grupo na lista (em minúsculas, sem acento).
# Só acrescente depois de confirmar com a equipe de rede.
APELIDOS = {
    "CAOA CHERY": {"original xangai": "simpar", "grande norte": "newchase"},
    "GAC": {"automob": "simpar"},
    "OMODA & JAECOO": {"kazza": "newchase"},
    "JETOUR": {"kazza motors": "newchase"},
    "LEAPMOTOR": {"via marconi": "isar", "original": "simpar"},
    "VOLVO": {},
    "GEELY": {}, "MG MOTOR": {}, "DENZA": {}, "ZEEKR": {},
}

# =====================================================================================================
# FERRAMENTAS COMUNS
# =====================================================================================================
HOJE = datetime.date.today()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"


LOGFILE = None


def log(msg=""):
    print(msg, flush=True)
    if LOGFILE:                                    # tudo que aparece na tela também vai para log.txt (para mandar a quem der suporte)
        try:
            with open(LOGFILE, "a", encoding="utf-8") as f:
                f.write(str(msg) + "\n")
        except Exception:
            pass


class ErroAcesso(RuntimeError):
    """Falha ao acessar um site, já explicada em português. 'transitorio' = vale tentar de novo."""
    def __init__(self, tipo, texto, transitorio=False):
        super().__init__(texto); self.tipo = tipo; self.transitorio = transitorio


def classificar(e, url=""):
    host = urllib.parse.urlparse(url).netloc or str(url)[:40]
    if isinstance(e, ErroAcesso):
        return e
    if isinstance(e, urllib.error.HTTPError):
        c = e.code
        if c == 401:
            return ErroAcesso("assinatura", f"{host} exige assinatura do navegador (HTTP 401). Use o arquivo manual (plano B do LEIA-ME).")
        if c == 403:
            return ErroAcesso("bloqueio", f"{host} bloqueou o acesso (HTTP 403): rede/VPN da empresa ou proteção do site. Teste fora da VPN ou use o arquivo manual.")
        if c == 404:
            return ErroAcesso("mudou", f"o endereço de {host} não existe mais (HTTP 404): o site provavelmente mudou. Avise quem mantém o programa ou use o plano B.")
        if c == 429 or c >= 500:
            return ErroAcesso("instavel", f"{host} está instável ou sobrecarregado (HTTP {c}).", True)
        return ErroAcesso("http", f"{host} respondeu HTTP {c}.")
    reason = getattr(e, "reason", e)
    txt = str(reason)
    if isinstance(reason, ssl.SSLError) or "CERTIFICATE_VERIFY_FAILED" in txt:
        return ErroAcesso("certificado", "erro de certificado: a rede da empresa intercepta conexões seguras. Rode 'pip install truststore' e tente de novo (o Iniciar_Rede.bat já tenta instalar).")
    if "Tunnel connection failed" in txt or "Proxy" in txt or "proxy" in txt:
        return ErroAcesso("proxy", f"o proxy/firewall da empresa bloqueou o acesso a {host}. Teste fora da rede corporativa ou use o arquivo manual.")
    if isinstance(reason, socket.gaierror) or "getaddrinfo" in txt or "Name or service" in txt:
        return ErroAcesso("dns", f"não consegui localizar {host} (sem internet, VPN ou DNS bloqueado).", True)
    if isinstance(e, (socket.timeout, TimeoutError)) or "timed out" in txt:
        return ErroAcesso("tempo", f"{host} demorou demais para responder.", True)
    return ErroAcesso("rede", f"falha de conexão com {host} ({txt[:100]}).", True)


def http(url, metodo="GET", corpo=None, headers=None, timeout=60, tentativas=3):
    h = {"User-Agent": UA, "Accept-Language": "pt-BR"}
    h.update(headers or {})
    ultimo = None
    for i in range(tentativas):
        espera = 0
        try:
            req = urllib.request.Request(url, data=corpo, method=metodo, headers=h)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode("utf-8", "ignore")
        except Exception as e:
            ultimo = classificar(e, url)
            if isinstance(e, urllib.error.HTTPError) and e.headers.get("Retry-After", "").isdigit():
                espera = min(int(e.headers["Retry-After"]), 30)
        if not ultimo.transitorio:
            break                                   # 401/403/404/certificado/proxy: repetir não adianta
        if i < tentativas - 1:
            time.sleep(max(espera, 2 * 2 ** i + random.random()))   # espera crescente: ~2 s, ~4 s...
    raise ultimo


def norm(s):
    """minúsculas, sem acento, só letras/números/espaço"""
    s = unicodedata.normalize("NFD", str(s if s is not None else "").lower()).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", s)).strip()


def digitos(s):
    return re.sub(r"\D", "", str(s if s is not None else ""))


def hav_arr(lat1, lng1, lat2, lng2):
    """distância em km; aceita arrays"""
    a, b, c, d = map(np.radians, [np.asarray(lat1, float), np.asarray(lng1, float), np.asarray(lat2, float), np.asarray(lng2, float)])
    x = np.sin((c - a) / 2) ** 2 + np.cos(a) * np.cos(c) * np.sin((d - b) / 2) ** 2
    return 6371 * 2 * np.arcsin(np.sqrt(x))


def num(x):
    r = re.search(r"-?\d+(?:\.\d+)?", str(x if x is not None else "").replace(",", "."))
    return float(r.group(0)) if r else float("nan")


def cols(df, nomes):
    """garante que o DataFrame tem estas colunas (vazias se faltarem)"""
    for n in nomes:
        if n not in df.columns:
            df[n] = ""
    return df


def ler_manual(nome):
    p = os.path.join("manual", nome)
    if os.path.exists(p):
        with open(p, encoding="utf-8-sig", errors="ignore") as f:
            return f.read()
    return None


# Arquivos manuais: o site bloqueia o programa, então uma pessoa salva o conteúdo em manual/ (veja o LEIA-ME).
MANUAIS_OBRIGATORIOS = {"GAC": "gac.json", "ZEEKR": "zeekr.json"}          # sem eles essas marcas não rodam
MANUAIS_RESERVA = {"VOLVO": "volvo.html", "JETOUR": "jetour.json", "CAOA_CHERY": "caoa_chery.html",
                   "GEELY": "geely.html", "MG": "mg.html"}                  # só servem de plano B se o site falhar
IDADE_MAX_MANUAL = 35                                                       # dias; acima disso o aviso aparece


def idade_manual(nome):
    """dias desde que manual/<nome> foi salvo (None se não existe)"""
    p = os.path.join("manual", nome)
    if not os.path.exists(p):
        return None
    return (datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getmtime(p))).days


def estado_manuais():
    """Situação dos arquivos manuais (usado no início da rodada e pelo painel)."""
    out = []
    for dic, obrig in ((MANUAIS_OBRIGATORIOS, True), (MANUAIS_RESERVA, False)):
        for chave, arq in dic.items():
            d = idade_manual(arq)
            if d is None:
                sit = "FALTA" if obrig else "não precisa (só se o site falhar)"
            elif d > IDADE_MAX_MANUAL:
                sit = f"velho ({d} dias)" if obrig else f"existe, {d} dias (plano B)"
            else:
                sit = f"em dia ({d} dias)" if obrig else f"existe, {d} dias (plano B)"
            out.append({"chave": chave, "arquivo": arq, "obrigatorio": obrig, "dias": d, "situacao": sit,
                        "precisa_atencao": obrig and (d is None or d > IDADE_MAX_MANUAL)})
    return out


def pagina(url, manual_nome, headers=None):
    """baixa a página; se falhar usa manual/<arquivo>; devolve (texto, origem)"""
    try:
        return http(url, headers=headers), "site"
    except Exception as e:
        t = ler_manual(manual_nome)
        if t is not None:
            return t, f"arquivo manual/{manual_nome}"
        raise RuntimeError(f"{str(e).rstrip('. ')}. Plano B: salve a página no navegador (Ctrl+S, 'somente HTML') como manual/{manual_nome}")


# =====================================================================================================
# 1) DOWNLOADS  (cada função devolve um DataFrame com as lojas)
# =====================================================================================================
def baixar_byd(**_):
    url = "https://us-site-api.byd.com/byd-api/us/dealer/getFindDealer?country=BR&type={}"
    campos = ["organizationCode", "organizationName", "parentCode", "status", "feature", "stateCode", "stateName", "cityName",
              "cityCode", "address", "zipCode", "phoneNumber", "mail", "lat", "lng"]
    lojas = {}
    for tipo in ("sales", "Service"):
        dados = json.loads(http(url.format(tipo), headers={"Accept": "application/json"}))
        if isinstance(dados, dict):
            dados = dados.get("data") or dados.get("list") or []
        for d in dados:
            cod = d.get("organizationCode")
            if cod and cod not in lojas:
                lojas[cod] = {c: d.get(c) for c in campos}
    if not lojas:
        raise RuntimeError("nenhuma loja BYD baixada")
    return pd.DataFrame(list(lojas.values()))


def baixar_gwm(byd=None, **_):
    """O site da GWM devolve as lojas perto de um CEP; consultamos vários CEPs e juntamos tudo."""
    url = "https://www.gwmmotors.com.br/pt/concessionarias/_jcr_content.brdealerlist.{}.json"
    capitais = ["01310100", "20040020", "30130010", "70040010", "40020000", "60025000", "50010000", "80010000", "90010000",
                "69010000", "66010000", "74003010", "65010000", "57020000", "59010000", "64000000", "58010000", "49010000",
                "79002000", "78005000", "88010000", "29010000", "76801000", "69900000", "68900000", "69301000", "77001000"]
    campos = ["dealerNumber", "document", "storeName", "storeCity", "storeState", "storeAddress", "storeZipCode", "storeContact", "storeContactWhatsapp"]
    fila = list(capitais)
    if byd is not None and "zipCode" in byd.columns:
        vistos = set()
        for cep in byd.zipCode.map(digitos):
            if len(cep) == 8 and cep[:4] not in vistos:
                vistos.add(cep[:4]); fila.append(cep)
    lojas, feitos, falhas, seguidas, ultimo = {}, set(), 0, 0, None
    while fila:
        if seguidas >= 8:                           # 8 consultas seguidas falharam: não adianta insistir nas centenas restantes
            if not lojas:
                raise ultimo if isinstance(ultimo, ErroAcesso) else RuntimeError(f"o site da GWM não respondeu ({ultimo})")
            log("    GWM: muitas falhas seguidas; parando com o que já foi lido")
            break
        cep = fila.pop(0)
        if cep in feitos:
            continue
        feitos.add(cep)
        try:
            j = json.loads(http(url.format(cep), headers={"Accept": "application/json"}, tentativas=1))
            dados = (j.get("result") or {}).get("data") or []
            seguidas = 0
        except Exception as e:
            falhas += 1; seguidas += 1; ultimo = e; dados = []
        for d in dados:
            k = f'{d.get("dealerNumber")}|{d.get("storeName")}|{d.get("storeZipCode")}'   # dealerNumber se repete; chave composta
            if k not in lojas:
                lojas[k] = {c: d.get(c) for c in campos}
                c2 = digitos(d.get("storeZipCode"))
                if len(c2) == 8 and c2 not in feitos:
                    fila.append(c2)
        if len(feitos) % 40 == 0:
            log(f"    GWM: {len(feitos)} CEPs consultados, {len(lojas)} lojas, fila {len(fila)}")
        time.sleep(0.25)
    if not lojas:
        raise RuntimeError(f"nenhuma loja GWM baixada ({falhas} de {len(feitos)} consultas falharam; veja a conexão)")
    df = pd.DataFrame(list(lojas.values()))
    if falhas > 0.25 * len(feitos):
        df.attrs["parcial"] = f"{falhas} de {len(feitos)} consultas ao site da GWM falharam"
    return df


def baixar_omoda_jaecoo(**_):
    base = "https://omodajaecoo.com.br"
    campos = ["id", "name", "filtro", "type", "status", "address", "complement", "district", "city", "state", "zip_code",
              "latitude", "longitude", "phone", "email", "site", "date_update", "date_deleted"]
    html_ = http(base + "/concessionarias")
    m = re.search(r'"buildId"\s*:\s*"([^"]+)"', html_)
    if not m:
        raise RuntimeError("não achei o buildId na página da Omoda & Jaecoo (o site pode ter mudado)")
    dados = json.loads(http(f"{base}/_next/data/{m.group(1)}/concessionarias.json"))

    def achar(o):
        if isinstance(o, dict):
            if isinstance(o.get("concessionaria"), list):
                return o["concessionaria"]
            for v in o.values():
                r = achar(v)
                if r:
                    return r
        return None
    lista = achar(dados)
    if not lista:
        raise RuntimeError("não achei a lista de concessionárias no JSON da Omoda & Jaecoo")

    def txt(v):
        if v is None:
            return ""
        return ", ".join(str(x) for x in v) if isinstance(v, (list, tuple)) else v
    return pd.DataFrame([{c: txt(d.get(c)) for c in campos} for d in lista])


def baixar_jetour(**_):
    url = "https://jetourbr.com/wp-admin/admin-ajax.php?action=jetour_lojas_near"
    try:
        dados = json.loads(http(url, headers={"Accept": "application/json, */*", "Referer": "https://jetourbr.com/concessionarias/"}, tentativas=1))
        origem = "site"
    except Exception as e:
        t = ler_manual("jetour.json")
        if t is None:
            raise RuntimeError(f"{str(e).rstrip('. ')}. Plano B: abra {url} no Chrome, salve como manual/jetour.json")
        dados, origem = json.loads(t), "arquivo manual/jetour.json"
    itens = (dados.get("data") or {}).get("items") or []
    if not itens:
        raise RuntimeError("nenhuma loja Jetour na resposta")
    linhas = []
    for it in itens:
        p = [x.strip() for x in str(it.get("label")).replace("–", "-").split(" - ")]
        linhas.append({"id": it.get("id"), "label": it.get("label"), "grupo_pelo_nome": p[0] if p else "",
                       "cidade_pelo_nome": p[1] if len(p) > 1 else "", "address": it.get("address"), "cep": it.get("cep"),
                       "dealership_id": it.get("dealership_id"), "servicos": it.get("servicos")})
    df = pd.DataFrame(linhas)
    df.attrs["origem"] = origem
    return df


def baixar_gac(**_):
    """A API da GAC exige uma assinatura que só o navegador gera. Por isso o caminho normal é o arquivo manual/gac.json."""
    t = ler_manual("gac.json")
    origem = "arquivo manual/gac.json"
    if t is None:
        corpo = json.dumps({"aliasName": ""}).encode("utf-8")
        h = {"Content-Type": "application/json", "Accept": "application/json", "Origin": "https://www.gacgroup.com",
             "Referer": "https://www.gacgroup.com/", "region": "BR", "locale": "pt_BR", "fnc-app-id": "fe-official",
             "fnc-requestid": str(uuid.uuid4()), "fnc-timestamp": str(int(time.time() * 1000))}
        try:
            t = http("https://br-www-api.gacgroup.com/gateway/v1/www/api/dealer/list-by-nearby", "POST", corpo, h, tentativas=1)
            origem = "site"
        except Exception as e:
            raise RuntimeError("a GAC exige assinatura do navegador. Faça: abra o site de concessionárias da GAC, F12 > Rede > "
                               "linha list-by-nearby > Resposta > copiar tudo, e salve como manual/gac.json")
    dados = json.loads(t)
    if not dados.get("success"):
        raise RuntimeError(f"a GAC recusou a consulta (code={dados.get('code')}; 4407 = assinatura inválida). Atualize o arquivo manual/gac.json (passo a passo no LEIA-ME).")
    lista = (dados.get("data") or {}).get("list") or []
    if not lista:
        raise RuntimeError("nenhuma loja GAC na resposta")
    linhas = []
    for d in lista:
        lat, lng = num(d.get("lat")), num(d.get("lng"))
        corrigida = ""
        if lat < -34.5:                       # o site tem lojas com latitude/longitude trocadas
            lat, lng, corrigida = lng, lat, "sim (estava trocada)"
        linhas.append({"id": d.get("id"), "code": d.get("code"), "name": d.get("name"), "cidade": d.get("secondLevelRegionName") or "",
                       "uf": d.get("firstLevelRegionName") or "", "address": d.get("address"), "lat": lat, "lng": lng,
                       "phone": d.get("phone"), "salesStatus": d.get("salesStatus"), "afterSaleStatus": d.get("afterSaleStatus"),
                       "vende": "sim" if d.get("salesStatus") == 1 else "nao (pos-venda)", "coordenada_corrigida": corrigida})
    df = pd.DataFrame(linhas)
    df.attrs["origem"] = origem
    return df


def baixar_zeekr(**_):
    """A API da Zeekr exige assinatura do navegador. O caminho é o arquivo manual/zeekr.json
    (resposta da linha .../store/list copiada do F12, aba Rede)."""
    t = ler_manual("zeekr.json")
    if t is None:
        raise RuntimeError("a Zeekr exige assinatura do navegador. Faça: abra o site de concessionárias da Zeekr, F12 > Rede > "
                           "linha 'list' (.../store/list) > Resposta > copiar tudo, e salve como manual/zeekr.json")
    dados = json.loads(t)
    itens = (dados.get("data") or {}).get("items") or []
    if not itens:
        raise RuntimeError("nenhuma loja Zeekr no arquivo manual/zeekr.json (confira se copiou a resposta inteira)")
    linhas = []
    for d in itens:
        a = d.get("storeAddress") or {}
        lat, lng = num(d.get("latitude")), num(d.get("longitude"))
        if lat < -34.5:                                  # o site tem lojas com latitude/longitude trocadas
            lat, lng = lng, lat
        rua = re.sub(r"\s+", " ", str(a.get("street") or "")).replace("Brazil Others", "").strip(" ,")
        uf = re.search(r"[-/ ,]\s*([A-Z]{2})\b[,\s]*(?:\d{2}\.?\d{3}-?\d{3})?\s*$", rua)
        tipo = d.get("storeType") or ""
        linhas.append({"storeCode": d.get("storeCode"), "name": d.get("storeName"), "tipo": tipo, "status": d.get("storeStatusName"),
                       "cidade": a.get("cityName") if a.get("cityName") not in (None, "", "Others") else "", "uf": uf.group(1) if uf else "",
                       "cep": a.get("postcode"), "address": rua, "lat": lat, "lng": lng,
                       "vende": "nao (pos-venda)" if tipo.lower().startswith("service") else "sim",
                       "valido_ate": str(d.get("serviceTimeEnd") or "")[:10]})
    df = pd.DataFrame(linhas)
    vazio = df.cidade.eq("")      # cidade ausente no site: tira "ZEEKR Space/House/Pop-Up" do nome
    df.loc[vazio, "cidade"] = df.loc[vazio, "name"].str.replace(r"(?i)^zeekr\s+(space|house|pop-up|service branch)\s*", "", regex=True).str.strip(" -")
    # mesma loja listada duas vezes (ex.: Pop-Up e House no mesmo endereço): fica uma só, preferindo a que não é Pop-Up
    df["_k"] = df.lat.round(2).astype(str) + "|" + df.lng.round(2).astype(str) + "|" + df.vende
    df["_p"] = df.tipo.str.lower().str.startswith("pop").astype(int)
    df = df.sort_values(["_p"]).drop_duplicates("_k").drop(columns=["_k", "_p"]).sort_values("storeCode").reset_index(drop=True)
    df.attrs["origem"] = "arquivo manual/zeekr.json"
    return df


def baixar_caoa_chery(**_):
    texto, origem = pagina("https://caoachery.com.br/concessionarias", "caoa_chery.html")
    t = html.unescape(texto)
    ini = re.compile(r'\{"id":(\d+),"name":"([^"]*)","short_name":"[^"]*","address_lat":"([^"]*)","address_lon":"([^"]*)","address":"([^"]*)"')

    def s(x):
        try:
            return json.loads('"' + x + '"')
        except Exception:
            return x
    lojas = {}
    pos = [(m.start(), m) for m in ini.finditer(t)]
    for k, (p, m) in enumerate(pos):
        bloco = t[p:(pos[k + 1][0] if k + 1 < len(pos) else p + 30000)]

        def pega(ch):
            r = re.search(r'"%s":"([^"]*)"' % ch, bloco)
            return s(r.group(1)) if r else ""
        nome = s(m.group(2))
        lojas[m.group(1)] = {"id": m.group(1), "name": nome, "lat": num(m.group(3)), "lng": num(m.group(4)), "address": s(m.group(5)),
                             "cep": pega("address_cep"), "cidade": pega("address_city"), "uf": pega("address_state"), "cnpj": pega("cnpj"),
                             "vende": "nao (oficina)" if re.search(r"oficina", nome, re.I) else "sim"}
    if not lojas:
        raise RuntimeError("não achei lojas na página da Caoa Chery (o site pode ter mudado). Plano B: manual/caoa_chery.html")
    df = pd.DataFrame(list(lojas.values()))
    df.attrs["origem"] = origem
    return df


def baixar_leapmotor(**_):
    url = ("https://dealer-service.k8s.fcalatam.com.br/dealerws/dealer?brandId=82&market=pt-BR&showBlocked=false&loadAll=true"
           "&dealerType=EA+-+CONCESSIONARIA+SEMIPLENA+%7C+EV+-+CONCESSIONARIA+SHOW+ROOM+%7C+CP+-+CONCESSIONARIA+COMPACTA"
           "+%7C+PL+-+CONCESSIONARIA+PLENA+%7C+SM+-+CONCESSIONARIA+SEMIPLENA+%7C+PV+-+CONCESSIONARIA+POS+VENDAS"
           "+%7C+SO+-+CONCESSIONARIA+POS+VENDAS")
    dados = json.loads(http(url, headers={"Accept": "application/json", "Origin": "https://www.leapmotor.com.br", "Referer": "https://www.leapmotor.com.br/"}))
    linhas = []
    for d in dados:
        tipo = (d.get("dealerType") or "")[:2]
        linhas.append({"dealerCode": d.get("dealerCode"), "name": d.get("fantasyName"), "razao_social": d.get("corporateName"),
                       "address": d.get("address"), "bairro": d.get("neighborhood"), "cidade": (d.get("city") or {}).get("name"),
                       "uf": d.get("state"), "cep": d.get("postalCode"), "lat": d.get("latitude"), "lng": d.get("longitude"),
                       "telefone": d.get("telephone"), "tipo": tipo,
                       "vende": "sim" if tipo in ("EA", "EV", "CP", "PL", "SM") else "nao (pos-venda)"})
    if not linhas:
        raise RuntimeError("nenhuma loja Leapmotor na resposta")
    return pd.DataFrame(linhas)


def baixar_denza(**_):
    url = "https://us-site-api.byd.com/tengshi-overseas-api/us/dealer/getUsDealer?brand=DENZA&country=BR"
    dados = json.loads(http(url, headers={"Accept": "application/json", "Origin": "https://www.denza.com", "Referer": "https://www.denza.com/br/concessionarias"}))
    if isinstance(dados, dict):
        dados = dados.get("data") or dados.get("list") or list(dados.values())
    linhas = []
    for d in dados:
        linhas.append({"organizationCode": d.get("organizationCode"), "name": d.get("organizationName"), "status": d.get("status"),
                       "feature": d.get("feature"), "cidade": d.get("cityName"), "uf": d.get("stateName"), "cep": d.get("zipCode"),
                       "address": d.get("address"), "lat": d.get("lat"), "lng": d.get("lng"), "telefone": d.get("phoneNumber"),
                       "vende": "nao (pos-venda)" if str(d.get("feature") or "").lower().startswith("serv") else "sim"})
    if not linhas:
        raise RuntimeError("nenhuma loja Denza na resposta")
    return pd.DataFrame(linhas)


def baixar_geely(**_):
    texto, origem = pagina("https://www.geelybrasil.com.br/contato", "geely.html")
    i = texto.find('name="unit"')
    if i < 0:
        raise RuntimeError("não achei a lista de unidades da Geely (o site pode ter mudado). Plano B: salve a página /contato como manual/geely.html")
    bloco = texto[i:texto.find("</select>", i)]
    linhas = []
    for o in [html.unescape(x) for x in re.findall(r'<option value="([^"]*)"', bloco) if x.strip()]:
        nome, _, local = o.rpartition(" - ")
        cidade, _, uf = local.partition("/")
        pos = bool(re.search(r"P[ÓO]S[- ]?VENDA|OFICINA", o, re.I))
        linhas.append({"name": nome.strip(), "cidade": cidade.strip(), "uf": uf.strip(), "vende": "nao (pos-venda)" if pos else "sim", "texto_original": o})
    df = pd.DataFrame(linhas)
    df.attrs["origem"] = origem
    return df


def baixar_mg(**_):
    texto, origem = pagina("https://mgmotoroficial.com.br/mg-care/concessionarias", "mg.html")
    itens = re.findall(r'class="_itemName_[^"]*">([^<]*)</span><span class="_itemAddress_[^"]*">([^<]*)</span>(?:<span class="_itemPhone_[^"]*">([^<]*)</span>)?', texto)
    if not itens:
        raise RuntimeError("não achei as concessionárias da MG (o site pode ter mudado). Plano B: manual/mg.html")
    linhas = []
    for nome, end, tel in itens:
        nome, end = html.unescape(nome).strip(), html.unescape(end).strip()
        k = end.rfind(" - ")
        uf, resto = end[k + 3:].strip(), end[:k]
        breve = "EM BREVE" in nome.upper()
        linhas.append({"name": re.sub(r"\s*\(EM BREVE\)", "", nome, flags=re.I).strip(), "cidade": resto[resto.rfind(",") + 1:].strip(), "uf": uf,
                       "address": end, "telefone": html.unescape(tel).strip(), "vende": "nao (em breve)" if breve else "sim", "texto_original": nome})
    df = pd.DataFrame(linhas)
    df.attrs["origem"] = origem
    return df


def baixar_volvo(**_):
    texto, origem = pagina("https://www.volvocars.com/br/dealers/concessionarios/", "volvo.html")
    t = texto.replace('\\"', '"')
    rx = re.compile(r'"addressLine1":"([^"]*)","addressLine2":"([^"]*)","city":"([^"]*)"[^}]*?"name":"([^"]*)","longitude":"([^"]*)","latitude":"([^"]*)"([\s\S]{0,900}?)"capabilities":\[([^\]]*)\]')
    vistos, linhas = set(), []
    for m in rx.finditer(t):
        a1, a2, cidade, nome, lng, lat, meio, caps = m.groups()
        if (nome, lng) in vistos:
            continue
        vistos.add((nome, lng))
        email = re.search(r'"generalContactEmail":"([^"]*)"', meio)
        tel = re.search(r'"main":"([^"]*)"', meio)
        linhas.append({"name": html.unescape(nome).strip(), "cidade": html.unescape(cidade).strip(), "address": html.unescape(a1 + ", " + a2).strip(),
                       "lat": num(lat), "lng": num(lng), "telefone": tel.group(1) if tel else "", "email": email.group(1) if email else "",
                       "capabilities": caps.replace('"', ""), "vende": "sim" if '"sales"' in caps else "nao (pos-venda)"})
    if not linhas:
        raise RuntimeError("não achei as concessionárias da Volvo (o site pode ter mudado). Plano B: manual/volvo.html")
    df = pd.DataFrame(linhas)
    df.attrs["origem"] = origem
    return df


# =====================================================================================================
# 2) COMPARAÇÕES  (cada função devolve um dicionário padrão; veja resultado_vazio)
# =====================================================================================================
def resultado_vazio():
    return {"n_site": 0, "n_lista": 0, "iguais": 0, "novas": [], "fechadas": [], "diverg": [], "outras": [], "notas": []}


def _coords(x):
    """latitude/longitude de uma linha do site, se existirem (senão vazio)"""
    try:
        for a, b in (("lat", "lng"), ("latitude", "longitude"), ("lat", "lon")):
            if a in x.index and b in x.index:
                la, lo = num(x[a]), num(x[b])
                if la == la and lo == lo:
                    return la, lo
    except Exception:
        pass
    return "", ""


def nova(loja, cidade="", uf="", endereco="", grupo="", detalhe="", x=None):
    la, lo = _coords(x) if x is not None else ("", "")
    return {"loja_no_site": loja, "cidade": cidade, "uf": uf, "endereco": endereco, "grupo_provavel": grupo, "detalhe": detalhe,
            "lat": la, "lng": lo}


def fechada(op, detalhe=""):
    return {"grupo": op["NOME DO GRUPO"], "cidade": op["CIDADE"], "uf": op["UF"], "endereco_na_lista": op["ENDEREÇO"], "detalhe": detalhe,
            "linha": op.get("_linha")}


def diverg(op, loja, end_site, km=None, detalhe=""):
    return {"grupo": op["NOME DO GRUPO"], "cidade": op["CIDADE"], "endereco_lista": op["ENDEREÇO"], "loja_no_site": loja,
            "endereco_no_site": end_site, "km": None if km is None else round(float(km), 1), "detalhe": detalhe}


# ---- BYD: grupo pelo domínio do e-mail ------------------------------------------------------------
def comparar_byd(B, lb, **_):
    r = resultado_vazio()
    B = cols(B.copy(), ["feature", "mail", "organizationName", "cityName", "stateName", "address"])
    B["lat"] = B.lat.map(num); B["lng"] = B.lng.map(num)
    B["_p"] = B.feature.notna().astype(int)
    B = B.sort_values("_p", ascending=False).reset_index(drop=True)
    usados, manter = set(), []
    for i, x in B.iterrows():                       # junta vendas + pós-venda da mesma loja física
        if i in usados:
            continue
        d = hav_arr(x.lat, x.lng, B.lat.values, B.lng.values)
        usados.update(int(j) for j in np.where(d < KM_MESMA_LOJA_BYD)[0])
        manter.append(i)
    U = B.loc[manter].reset_index(drop=True)

    def grupo_email(m):
        m = str(m).lower() if isinstance(m, str) else ""
        return next((v for k, v in DOMINIOS_BYD.items() if k in m), None)
    U["grupo_prov"] = U.mail.map(grupo_email)
    lb = lb.reset_index(drop=True)
    U["km_lista"] = [float(hav_arr(x.lat, x.lng, lb.LATITUDE.values, lb.LONGITUDE.values).min()) if len(lb) else 999 for _, x in U.iterrows()]
    U["na_lista"] = U.km_lista < KM_BYD
    fora = U[U.grupo_prov.notna() & ~U.na_lista]
    for _, x in fora.iterrows():
        r["novas"].append(nova(x.organizationName, x.cityName, x.stateName, x.address, x.grupo_prov, f"{x.km_lista:.1f} km da operação mais próxima; grupo pelo e-mail ({x.mail})", x=x))
    sem = 0
    for _, op in lb.iterrows():
        d = hav_arr(op.LATITUDE, op.LONGITUDE, U.lat.values, U.lng.values)
        if len(d) == 0 or d.min() >= KM_BYD:
            r["fechadas"].append(fechada(op, f"nenhuma loja BYD a menos de {KM_BYD:.0f} km")); sem += 1
    for _, x in U[U.grupo_prov.isna() & ~U.na_lista].iterrows():
        r["outras"].append(nova(x.organizationName, x.cityName, x.stateName, x.address, "", x.mail if isinstance(x.mail, str) else ""))
    r.update(n_site=len(U), n_lista=len(lb), iguais=len(lb) - sem)
    r["notas"].append(f"{len(B)} registros (vendas+pós) viraram {len(U)} lojas físicas. Grupo deduzido pelo e-mail (inferência).")
    return r


# ---- GWM: grupo pelo CNPJ -------------------------------------------------------------------------
def comparar_gwm(G, lg, **_):
    r = resultado_vazio()
    G = cols(G.copy(), ["document", "storeZipCode", "storeCity", "storeState", "storeName", "storeAddress"]).fillna("")
    lg = lg.reset_index(drop=True).copy()
    if "GWM" in MARCAS_SEM_SHOPPING:
        pad = re.compile(PADRAO_SHOPPING, re.I)
        sh_s = (G.storeName + " " + G.storeAddress).map(lambda t: bool(pad.search(t)))
        sh_l = lg["ENDEREÇO"].map(lambda t: bool(pad.search(str(t))))
        r["notas"].append(f"Lojas em shopping/outlet ignoradas: {int(sh_s.sum())} no site e {int(sh_l.sum())} na lista (só dealers com ponto próprio).")
        G, lg = G[~sh_s].copy(), lg[~sh_l].reset_index(drop=True)
    G["raiz"] = G.document.map(lambda x: digitos(x)[:8])
    G["cep"] = G.storeZipCode.map(digitos); G["cid"] = G.storeCity.map(lambda x: norm(x).replace(" ", ""))
    lg["cep"] = lg.CEP.map(digitos); lg["cid"] = lg.CIDADE.map(lambda x: norm(x).replace(" ", ""))
    votos = {}
    for _, op in lg.iterrows():
        m = G[G.cep == op.cep]
        if m.empty:                              # sem CEP igual: usa a cidade, só se todas as lojas dela forem do mesmo CNPJ
            m = G[(G.cid == op.cid) & (G.storeState == op.UF)]
            if m.raiz.nunique() != 1:
                continue
        for raiz in set(m.raiz) - {""}:
            votos.setdefault(raiz, []).append(op["NOME DO GRUPO"])
    mapa = {k: max(set(v), key=v.count) for k, v in votos.items()}
    G["grupo"] = G.raiz.map(mapa)
    lg["chave"] = lg["NOME DO GRUPO"] + "|" + lg.cid + "|" + lg.UF
    t = G[G.grupo.notna()].copy()
    t["chave"] = t.grupo + "|" + t.cid + "|" + t.storeState
    api_n, lis_n = t.groupby("chave").size(), lg.groupby("chave").size()
    iguais = 0
    for k in sorted(set(api_n.index) | set(lis_n.index)):
        a, b = int(api_n.get(k, 0)), int(lis_n.get(k, 0))
        iguais += min(a, b)
        grupo, _, uf = k.split("|")[0], None, k.split("|")[-1]
        cidade = next(iter(t[t.chave == k].storeCity), None) or next(iter(lg[lg.chave == k].CIDADE), k)
        if a > b:
            lojas = t[t.chave == k]
            r["novas"].append(nova("; ".join(lojas.storeName), cidade, uf, "; ".join(lojas.storeAddress), grupo, f"site tem {a} loja(s), a lista tem {b}"))
        elif a < b:
            ops = lg[lg.chave == k]
            for j in range(a, b):                      # uma linha para cada operação sem correspondência
                r["fechadas"].append(fechada(ops.iloc[j], f"lista tem {b} operação(ões) do grupo na cidade, o site tem {a}"))
    for raiz, g in G[G.grupo.isna()].groupby("raiz"):
        r["outras"].append(nova(g.storeName.iloc[0], ", ".join(sorted(set(g.storeCity))[:6]), "", "", "", f"CNPJ {raiz}: {len(g)} loja(s)"))
    r.update(n_site=len(G), n_lista=len(lg), iguais=iguais)
    r["notas"].append(f"Grupo identificado pelo CNPJ (confiável). {len(mapa)} CNPJs ligados a grupos da lista.")
    return r


# ---- Omoda & Jaecoo: grupo pelo prefixo do nome, aprendido por distância ---------------------------
def comparar_omoda_jaecoo(A, lo, **_):
    r = resultado_vazio()
    A = cols(A.copy(), ["name", "city", "state", "address"]).fillna("")
    A["lat"] = A.latitude.map(num); A["lng"] = A.longitude.map(num)
    lo = lo.reset_index(drop=True)
    ap = APELIDOS["OMODA & JAECOO"]
    A["palavras"] = A.name.map(lambda n: norm(n).split())
    A["prefixo"] = A.palavras.map(lambda p: ap.get(p[0], p[0]) if p else "")
    mapa, achadas = {}, {}
    for i, op in lo.iterrows():
        d = hav_arr(op.LATITUDE, op.LONGITUDE, A.lat.values, A.lng.values)
        j = int(np.nanargmin(d))
        if d[j] < KM_PADRAO:
            mapa[A.prefixo[j]] = op["NOME DO GRUPO"]; achadas[i] = j

    def grupo(x):
        for p in [x.prefixo] + x.palavras:
            p = ap.get(p, p)
            if p in mapa:
                return mapa[p]
        return None
    A["grupo"] = A.apply(grupo, axis=1)
    usadas = set(achadas.values())
    for _, x in A[A.grupo.notna() & ~A.index.isin(usadas)].iterrows():
        r["novas"].append(nova(x["name"], x.city, x.state, x.address, x.grupo, "grupo pelo nome da loja (inferência)", x=x))
    for i, op in lo.iterrows():
        if i not in achadas:
            r["fechadas"].append(fechada(op, f"nenhuma loja a menos de {KM_PADRAO:.0f} km"))
    for _, x in A[A.grupo.isna()].iterrows():
        r["outras"].append(nova(x["name"], x.city, x.state, x.address))
    r.update(n_site=len(A), n_lista=len(lo), iguais=len(achadas))
    r["notas"].append("Grupo deduzido pelo nome da loja (inferência).")
    return r


# ---- Jetour: grupo + cidade + número do endereço ---------------------------------------------------
def comparar_jetour(S, lj, **_):
    r = resultado_vazio()
    S = cols(S.copy(), ["grupo_pelo_nome", "cidade_pelo_nome", "address", "label", "servicos"]).fillna("")
    lj = lj.reset_index(drop=True).copy()
    ap = APELIDOS["JETOUR"]

    def numero(e):
        m = re.search(r"\b(\d{1,6})\b", str(e))
        return m.group(1) if m else ""
    S["g"] = S.grupo_pelo_nome.map(lambda x: ap.get(norm(x), norm(x)))
    S["c"] = S.cidade_pelo_nome.map(norm); S["n"] = S.address.map(numero)
    lj["g"] = lj["NOME DO GRUPO"].map(norm); lj["c"] = lj.CIDADE.map(norm); lj["n"] = lj["ENDEREÇO"].map(numero)
    grupos = set(lj.g)
    usadas, ok = set(), 0
    for _, op in lj.iterrows():
        m = S[(S.g == op.g) & (S.c == op.c) & (S.n == op.n) & ~S.index.isin(usadas)]
        if len(m):
            usadas.add(m.index[0]); ok += 1; continue
        m = S[(S.g == op.g) & (S.c == op.c) & ~S.index.isin(usadas)]
        if len(m):
            usadas.add(m.index[0]); r["diverg"].append(diverg(op, m.iloc[0].label, m.iloc[0].address, None, "mesmo grupo e cidade, número diferente")); continue
        r["fechadas"].append(fechada(op, "sem loja do grupo nessa cidade no site"))
    for _, x in S[S.g.isin(grupos) & ~S.index.isin(usadas)].iterrows():
        r["novas"].append(nova(x.label, x.cidade_pelo_nome, "", x.address, x.grupo_pelo_nome, f"serviços: {x.servicos}", x=x))
    for _, x in S[~S.g.isin(grupos)].iterrows():
        r["outras"].append(nova(x.label, x.cidade_pelo_nome, "", x.address))
    r.update(n_site=len(S), n_lista=len(lj), iguais=ok)
    r["notas"].append("O site da Jetour devolve só parte das lojas (60 de 69 na última conferência); possíveis 'fechadas' podem ser lojas que o site não mostra.")
    return r


# ---- Distância (GAC, Caoa Chery, Leapmotor, Volvo, Denza) -----------------------------------------
def prefixo_gac(nome, cid_site="", cid_lista=""):
    t = norm(nome).split()
    return t[1] if len(t) > 1 and t[0] == "gac" else (t[0] if t else "")


def prefixo_chery(nome, cid_site="", cid_lista=""):
    return norm(str(nome).split("|", 1)[-1].split(" - ")[0])


def prefixo_volvo(nome, cid_site="", cid_lista=""):
    return norm(str(nome).split(" - ")[0])


def prefixo_leap(nome, cid_site="", cid_lista=""):
    t = [w for w in norm(nome).split() if w not in {"pv", "lpm", "oficina"}]
    cid = set((norm(cid_site) + " " + norm(cid_lista)).split())
    while len(t) > 1 and t[-1] in cid:
        t.pop()
    return " ".join(t)


def prefixo_nenhum(nome, cid_site="", cid_lista=""):
    return ""


def comparar_distancia(S, lm, marca, prefixo, km=KM_PADRAO, **_):
    r = resultado_vazio()
    S = cols(S.copy(), ["name", "cidade", "uf", "address", "vende"]).fillna("")
    S["lat"] = S.lat.map(num); S["lng"] = S.lng.map(num)
    S["vende_b"] = S.vende.astype(str).str.startswith("sim") | (S.vende == "")
    lm = lm.reset_index(drop=True).copy()
    lm["g"] = lm["NOME DO GRUPO"].map(norm); lm["c"] = lm.CIDADE.map(norm)
    S["c"] = S.cidade.map(norm)
    V = S[S.vende_b]
    usadas, ok, pref2g, sem = {}, [], {}, []
    for i, op in lm.iterrows():
        cand = V[~V.index.isin(usadas)]
        d = pd.Series(hav_arr(op.LATITUDE, op.LONGITUDE, cand.lat.values, cand.lng.values), index=cand.index) if len(cand) else pd.Series(dtype=float)
        if len(d) and np.nanmin(d.values) <= km:
            j = d.idxmin(); usadas[j] = op.g; ok.append(i)
            p = prefixo(S.loc[j, "name"], S.loc[j, "cidade"], op.CIDADE)
            if p:
                pref2g[p] = op.g
        else:
            sem.append(i)
    for k, v in APELIDOS.get(marca, {}).items():
        pref2g[k] = v
    S["p"] = [prefixo(n, c, "") for n, c in zip(S["name"], S.cidade)]
    S["g"] = S.p.map(lambda p: pref2g.get(p, ""))
    grupos = set(lm.g)
    ainda = []
    for i in sem:                                  # sem loja a menos de km: tenta a mesma cidade (endereço/coordenada divergente)
        op = lm.loc[i]
        m = S[S.vende_b & ~S.index.isin(usadas) & (S.c == op.c)]
        mg = m[m.g == op.g]
        escolha = mg if len(mg) else (m[m.g == ""] if (len(m[m.g == ""]) == 1) else m.iloc[0:0])
        if len(escolha):
            j = escolha.index[0]; usadas[j] = op.g
            dk = float(hav_arr(op.LATITUDE, op.LONGITUDE, S.loc[j, "lat"], S.loc[j, "lng"]))
            r["diverg"].append(diverg(op, S.loc[j, "name"], S.loc[j, "address"], dk,
                                      "mesma cidade, a mais de %.0f km da lista (endereço novo ou coordenada imprecisa)" % km))
            ok.append(i)
            if S.loc[j, "g"] == "" and S.loc[j, "p"]:
                pref2g[S.loc[j, "p"]] = op.g
        else:
            ainda.append(i)
    S["g"] = S.p.map(lambda p: pref2g.get(p, ""))
    for _, x in S[S.vende_b & S.g.isin(grupos) & ~S.index.isin(usadas)].iterrows():
        r["novas"].append(nova(x["name"], x.cidade, x.uf, x.address, x.g, x=x))
    for i in ainda:
        r["fechadas"].append(fechada(lm.loc[i], f"nenhuma loja a menos de {km:.0f} km nem na mesma cidade"))
    for _, x in S[S.vende_b & ~S.g.isin(grupos) & ~S.index.isin(usadas)].iterrows():
        r["outras"].append(nova(x["name"], x.cidade, x.uf, x.address))
    posv = S[~S.vende_b & S.g.isin(grupos)]
    r.update(n_site=len(S), n_lista=len(lm), iguais=len(ok) - len(r["diverg"]))
    n_nao_vende = int((~S.vende_b).sum())
    if n_nao_vende:
        r["notas"].append(f"{n_nao_vende} unidade(s) do site são só pós-venda/oficina/'em breve' e não entram como loja (dessas, {len(posv)} de grupos da lista).")
    if marca == "ZEEKR":
        r["notas"].append("Zeekr: o site não informa o grupo; lojas fora da lista aparecem em 'Outros grupos' (grupo a confirmar). Pop-Up = loja provisória (veja 'valido_ate' no CSV).")
    elif marca == "DENZA":
        r["notas"].append("Denza: o site não informa o grupo; lojas fora da lista aparecem em 'Outros grupos' (grupo a confirmar). 'In interim' = loja provisória.")
    else:
        r["notas"].append("Grupo da loja do site deduzido pelo nome (inferência).")
    return r


# ---- Nome + cidade (Geely, MG: o site não informa coordenadas) -----------------------------
def comparar_nome_cidade(S, lm, marca, grupos_todos, **_):
    r = resultado_vazio()
    S = cols(S.copy(), ["name", "cidade", "uf", "address", "vende"]).fillna("")
    lm = lm.reset_index(drop=True).copy()
    lm["g"] = lm["NOME DO GRUPO"].map(norm); lm["c"] = lm.CIDADE.map(norm)
    ap = APELIDOS.get(marca, {})
    gs = sorted(grupos_todos, key=len, reverse=True)

    def grupo(nome):
        n = " " + norm(nome) + " "
        for k, v in ap.items():
            if f" {k} " in n:
                return v
        return next((g for g in gs if g and f" {g} " in n), "")
    S["g"] = S["name"].map(grupo); S["c"] = S.cidade.map(norm)
    S["vende_b"] = S.vende.astype(str).str.startswith("sim") | (S.vende == "")
    V = S[S.vende_b]
    gl = set(lm.g)
    cont_l = lm.groupby(["g", "c"]).size().rename("lista")
    cont_s = V[V.g.isin(gl)].groupby(["g", "c"]).size().rename("site")
    comp = pd.concat([cont_l, cont_s], axis=1).fillna(0).astype(int).reset_index()
    iguais = 0
    for _, x in comp.iterrows():
        iguais += min(x.lista, x.site)
        if x.site > x.lista:
            for _, l in V[(V.g == x.g) & (V.c == x.c)].iterrows():
                r["novas"].append(nova(l["name"], l.cidade, l.uf, l.address, x.g, f"site tem {x.site}, lista tem {x.lista}"))
        elif x.site < x.lista:
            ops = lm[(lm.g == x.g) & (lm.c == x.c)]
            em_breve = S[(S.g == x.g) & (S.c == x.c) & ~S.vende_b]
            det = f"lista tem {x.lista}, site tem {x.site}"
            if len(em_breve):
                det += " (o site mostra essa unidade como 'em breve'/pós-venda)"
            for j in range(int(x.site), int(x.lista)):
                r["fechadas"].append(fechada(ops.iloc[j], det))
    for _, l in V[~V.g.isin(gl)].iterrows():
        r["outras"].append(nova(l["name"], l.cidade, l.uf, l.address))
    r.update(n_site=len(S), n_lista=len(lm), iguais=int(iguais))
    r["notas"].append("Sem endereço/coordenada no site: comparação por grupo + cidade (contagem).")
    return r


# =====================================================================================================
# 3) CADASTRO DAS MARCAS
# =====================================================================================================
# chave, nome da marca na coluna MARCA da lista, função de download, função de comparação (recebe S, ops da marca)
def _dist(marca, pref, km=KM_PADRAO):
    return lambda S, lm, **kw: comparar_distancia(S, lm, marca, pref, km)


def _nc(marca):
    return lambda S, lm, grupos_todos, **kw: comparar_nome_cidade(S, lm, marca, grupos_todos)


MARCAS = [
    ("BYD", "BYD", baixar_byd, comparar_byd),
    ("GWM", "GWM", baixar_gwm, comparar_gwm),
    ("OMODA_JAECOO", "OMODA & JAECOO", baixar_omoda_jaecoo, comparar_omoda_jaecoo),
    ("JETOUR", "JETOUR", baixar_jetour, comparar_jetour),
    ("GAC", "GAC", baixar_gac, _dist("GAC", prefixo_gac)),
    ("CAOA_CHERY", "CAOA CHERY", baixar_caoa_chery, _dist("CAOA CHERY", prefixo_chery)),
    ("LEAPMOTOR", "LEAPMOTOR", baixar_leapmotor, _dist("LEAPMOTOR", prefixo_leap)),
    ("DENZA", "DENZA", baixar_denza, _dist("DENZA", prefixo_nenhum)),
    ("VOLVO", "VOLVO", baixar_volvo, _dist("VOLVO", prefixo_volvo)),
    ("GEELY", "GEELY", baixar_geely, _nc("GEELY")),
    ("MG", "MG MOTOR", baixar_mg, _nc("MG MOTOR")),
    ("ZEEKR", "ZEEKR", baixar_zeekr, _dist("ZEEKR", prefixo_nenhum)),
]

# --- Marcas cadastradas em arquivo (marcas_extras.csv): para sites com uma API JSON simples, sem programar ---
def _caminho(obj, caminho):
    """acessa 'data.list' / 'dealer.city.name' dentro de um JSON (listas: use o número, ex.: 'items.0.nome')"""
    for p in [x for x in str(caminho or "").split(".") if x]:
        if isinstance(obj, list):
            obj = obj[int(p)] if p.isdigit() and int(p) < len(obj) else None
        elif isinstance(obj, dict):
            obj = obj.get(p)
        else:
            return None
        if obj is None:
            return None
    return obj


def baixar_generico(cfg):
    def f(**_):
        origem = "site"
        try:
            t = http(cfg["url"], headers={"Accept": "application/json"})
        except Exception as e:
            arq = cfg.get("arquivo_manual", "")
            t = ler_manual(arq) if arq else None
            if t is None:
                raise RuntimeError(str(e).rstrip(". ") + (f". Plano B: salve a resposta da API como manual/{arq}" if arq else ""))
            origem = f"arquivo manual/{arq}"
        try:
            dados = json.loads(t)
        except Exception:
            raise ErroAcesso("mudou", f"a resposta de {cfg['chave']} não é JSON (o site mudou ou pediu verificação de robô).")
        lista = _caminho(dados, cfg.get("caminho_lista"))
        if not isinstance(lista, list) or not lista:
            raise ErroAcesso("mudou", f"não achei a lista de lojas em '{cfg.get('caminho_lista')}' na resposta de {cfg['chave']}: o formato do site mudou.")
        campos = {k: cfg.get(k, "") for k in ("name", "cidade", "uf", "address", "lat", "lng", "vende")}
        linhas = [{k: ("" if not c else _caminho(d, c)) for k, c in campos.items()} for d in lista]
        df = pd.DataFrame(linhas)
        df.attrs["origem"] = origem
        return df
    return f


def carregar_marcas_extras(arq="marcas_extras.csv"):
    if not os.path.exists(arq):
        return
    try:
        T = pd.read_csv(arq, sep=";", dtype=str, encoding="utf-8-sig").fillna("")
    except Exception as e:
        log(f"AVISO: não consegui ler {arq}: {e}")
        return
    ja = {m[0] for m in MARCAS}
    for _, r in T.iterrows():
        cfg = {k: str(v).strip() for k, v in r.items()}
        if not cfg.get("chave") or cfg["chave"].startswith("#"):
            continue
        falta = [k for k in ("chave", "nome_na_lista", "url", "caminho_lista", "lat", "lng") if not cfg.get(k)]
        if falta or cfg["chave"] in ja:
            log(f"AVISO: linha '{cfg.get('chave')}' de {arq} ignorada ({'faltam ' + ', '.join(falta) if falta else 'chave já existe'}).")
            continue
        try:
            km = float(cfg.get("km") or KM_PADRAO)
        except ValueError:
            km = KM_PADRAO
        comp = _nc(cfg["nome_na_lista"]) if cfg.get("metodo", "distancia") == "nome_cidade" else _dist(cfg["nome_na_lista"], prefixo_nenhum, km)
        MARCAS.append((cfg["chave"], cfg["nome_na_lista"], baixar_generico(cfg), comp))
        APELIDOS.setdefault(cfg["nome_na_lista"], {})
        if cfg.get("esperado", "").isdigit():
            ESPERADO[cfg["chave"]] = int(cfg["esperado"])
        if cfg.get("arquivo_manual"):
            MANUAIS_RESERVA[cfg["chave"]] = cfg["arquivo_manual"]


carregar_marcas_extras()
SEM_COLETA = {}   # marcas da lista que não conseguimos automatizar (aparecem no resumo como "não automatizada")


# =====================================================================================================
# 4) EXECUÇÃO
# =====================================================================================================
def achar_planilha(arg):
    if arg:
        return arg
    for f in sorted(glob.glob("*.xlsx"), key=os.path.getmtime, reverse=True):
        if os.path.basename(f).startswith(("~$", "Comparacao_Rede_")):
            continue
        try:
            if ABA_LISTA in pd.ExcelFile(f).sheet_names:
                return f
        except Exception:
            pass
    raise SystemExit(f'Não encontrei nenhuma planilha com a aba "{ABA_LISTA}" nesta pasta. Coloque-a aqui ou use --lista arquivo.xlsx')


def csv_mais_recente(pasta, chave):
    arqs = sorted(glob.glob(os.path.join(pasta, f"lojas_{chave}_*.csv")))
    return arqs[-1] if arqs else None


def auto_largura(ws):
    from openpyxl.styles import Font, PatternFill
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
    for col in ws.columns:
        tam = max((len(str(c.value)) if c.value is not None else 0) for c in col[:200])
        ws.column_dimensions[col[0].column_letter].width = min(max(10, tam + 2), 60)
    ws.freeze_panes = "A2"


HIST_DIR = "historico"
CACHE_DIR = "cache"          # última leitura completa de cada marca (usada se o download falhar)


def baixar_com_retentativa(baixar, **kw):
    """Tenta o download; se a falha for temporária (rede, site instável), espera 20 s e tenta mais uma vez."""
    try:
        return baixar(**kw)
    except Exception as e:
        msg = str(e)
        definitivo = isinstance(e, ErroAcesso) and not e.transitorio
        if definitivo or re.search(r"HTTP 40|certificado|assinatura|manual/|bloque", msg):
            raise
        log(f"  falha temporária ({msg.splitlines()[0][:120]}); aguardando 20 s para tentar de novo...")
        time.sleep(20)
        return baixar(**kw)


TESTES_CONEXAO = [
    ("BYD", "https://us-site-api.byd.com/byd-api/us/dealer/getFindDealer?country=BR&type=sales"),
    ("GWM", "https://www.gwmmotors.com.br/pt/concessionarias/_jcr_content.brdealerlist.01310100.json"),
    ("OMODA & JAECOO", "https://omodajaecoo.com.br/concessionarias"),
    ("JETOUR", "https://jetourbr.com/wp-admin/admin-ajax.php?action=jetour_lojas_near"),
    ("GAC", "https://br-www-api.gacgroup.com/gateway/v1/www/api/dealer/list-by-nearby"),
    ("CAOA CHERY", "https://caoachery.com.br/concessionarias"),
    ("LEAPMOTOR", "https://dealer-service.k8s.fcalatam.com.br/dealerws/dealer?brandId=82&market=pt-BR&showBlocked=false&loadAll=true"),
    ("DENZA", "https://us-site-api.byd.com/tengshi-overseas-api/us/dealer/getUsDealer?brand=DENZA&country=BR"),
    ("VOLVO", "https://www.volvocars.com/br/dealers/concessionarios/"),
    ("GEELY", "https://www.geelybrasil.com.br/contato"),
    ("MG MOTOR", "https://mgmotoroficial.com.br/mg-care/concessionarias"),
]


def diagnostico_conexao():
    """Testa se o computador alcança o site de cada marca. Qualquer resposta do servidor (mesmo 401/403/405) conta como 'alcançou'."""
    log("Testando a conexão com os sites das marcas (até 15 s cada)...")
    linhas = []
    for marca, url in TESTES_CONEXAO:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=15) as r:
                res, ok = f"OK (HTTP {r.status})", True
        except urllib.error.HTTPError as e:
            c = classificar(e, url)
            ok = c.tipo in ("assinatura", "http", "instavel") or e.code == 405
            res = f"alcançou o servidor (HTTP {e.code})" + ("" if ok else f" - {c}")
        except Exception as e:
            c = classificar(e, url); ok = False; res = f"SEM ACESSO - {c}"
        linhas.append({"Marca": marca, "Resultado": res, "Acesso": "ok" if ok else "PROBLEMA"})
        log(f"  {'✔' if ok else '✘'} {marca}: {res}")
    n = sum(1 for x in linhas if x["Acesso"] != "ok")
    log("\nTodos os sites responderam." if not n else f"\n{n} site(s) com problema. Se forem vários, o bloqueio costuma ser da rede/VPN da empresa: tente fora dela.")
    return linhas
UFS = set("AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO".split())


def qualidade_lista(L):
    """Procura problemas na própria lista da rede (coordenada ausente/fora do Brasil, UF inválida, duplicidade, UF x região)."""
    probs = []

    def add(i, row, texto):
        probs.append({"linha_na_planilha": int(i) + 2, "marca": row["MARCA"], "grupo": row["NOME DO GRUPO"], "cidade": row["CIDADE"],
                      "uf": row["UF"], "endereco": row["ENDEREÇO"], "problema": texto})
    ok = L.LATITUDE.notna() & L.LONGITUDE.notna()
    med = L[ok].groupby(L.UF.str.upper())[["LATITUDE", "LONGITUDE"]].median()
    vistos = {}
    for i, r in L.iterrows():
        if not ok[i]:
            add(i, r, "sem latitude/longitude (a comparação por distância não funciona para esta operação)"); continue
        if not (-34.5 <= r.LATITUDE <= 6 and -74.5 <= r.LONGITUDE <= -34):
            add(i, r, "coordenada fora do Brasil (latitude/longitude trocadas ou com erro)")
        if str(r.UF).upper() not in UFS:
            add(i, r, f"UF inválida: '{r.UF}'")
        elif str(r.UF).upper() in med.index and len(L[L.UF.str.upper() == str(r.UF).upper()]) >= 3:
            m = med.loc[str(r.UF).upper()]
            d = float(hav_arr(r.LATITUDE, r.LONGITUDE, m.LATITUDE, m.LONGITUDE))
            if d > 900:
                add(i, r, f"a coordenada está a {d:.0f} km das outras operações da UF {r.UF} (UF ou coordenada errada?)")
        k = (r["MARCA"], round(r.LATITUDE, 4), round(r.LONGITUDE, 4))
        if k in vistos:
            add(i, r, f"mesma marca e coordenada da linha {vistos[k] + 2} (possível duplicidade)")
        else:
            vistos[k] = i
        if not str(r["ENDEREÇO"]).strip():
            add(i, r, "endereço vazio")
    return pd.DataFrame(probs, columns=["linha_na_planilha", "marca", "grupo", "cidade", "uf", "endereco", "problema"])


def carregar_historico(hoje_nome):
    arqs = [a for a in sorted(glob.glob(os.path.join(HIST_DIR, "*.json"))) if os.path.basename(a) != hoje_nome]
    if not arqs:
        return {}
    try:
        with open(arqs[-1], encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def chave_item(tipo, marca, cidade, grupo):
    return "|".join([tipo, marca, norm(cidade), norm(grupo)])


def salvar_xlsx(arq, abas):
    def gravar(destino):
        with pd.ExcelWriter(destino, engine="openpyxl") as w:
            for nome, df in abas.items():
                df.to_excel(w, sheet_name=nome, index=False)
            for ws in w.book.worksheets:
                auto_largura(ws)
    try:
        gravar(arq)
        return arq
    except PermissionError:                        # arquivo aberto no Excel: grava com outro nome
        alt = arq.replace(".xlsx", f"_{datetime.datetime.now():%H%M%S}.xlsx")
        log(f"  (a planilha {os.path.basename(arq)} está aberta no Excel; gravando como {os.path.basename(alt)})")
        gravar(alt)
        return alt


UF_POR_NOME = {"acre": "AC", "alagoas": "AL", "amapa": "AP", "amazonas": "AM", "bahia": "BA", "ceara": "CE", "distrito federal": "DF",
               "espirito santo": "ES", "goias": "GO", "maranhao": "MA", "mato grosso": "MT", "mato grosso do sul": "MS", "minas gerais": "MG",
               "para": "PA", "paraiba": "PB", "parana": "PR", "pernambuco": "PE", "piaui": "PI", "rio de janeiro": "RJ",
               "rio grande do norte": "RN", "rio grande do sul": "RS", "rondonia": "RO", "roraima": "RR", "santa catarina": "SC",
               "sao paulo": "SP", "sergipe": "SE", "tocantins": "TO"}
NOTA_SAIDA = {"JETOUR": " ATENÇÃO: o site da Jetour não mostra todas as lojas; confirme antes de considerar fechada."}
COL_STATUS = "STATUS (PROGRAMA)"
COR_VERDE, COR_VERMELHO = "FFC6EFCE", "FFFFC7CE"


def _xml_texto(t):
    from xml.sax.saxutils import escape
    return escape(str(t), {'"': "&quot;"})


def _nome_saida(plan):
    """REP_261030_Grupos...xlsx -> REP_<hoje AAMMDD>_Grupos...xlsx (troca só a data); outro nome: acrescenta _COMPARADO_data"""
    base = os.path.basename(plan)
    m = re.match(r"^(.*?)(\d{6})(_.*\.xlsx)$", base, re.I)
    if m:
        return f"{m.group(1)}{datetime.date.today():%y%m%d}{m.group(3)}"
    return re.sub(r"\.xlsx$", "", base, flags=re.I) + f"_COMPARADO_{datetime.date.today():%Y-%m-%d}.xlsx"


def gerar_lista_comparada(plan, pasta, novas, saidas):
    """Copia a planilha da rede e, só na aba da lista: pinta de vermelho as operações que saíram, acrescenta em verde as lojas novas e
    cria a coluna STATUS (PROGRAMA). Mexe direto no XML do .xlsx para não perder imagens, formatações e fórmulas das outras abas."""
    import zipfile
    zin = zipfile.ZipFile(plan)
    wb = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(ABA_LISTA.replace("&", "&amp;")), wb)
    if not m:
        raise RuntimeError(f"aba '{ABA_LISTA}' não encontrada no arquivo")
    alvo = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % m.group(1), rels).group(0)
    caminho = "xl/" + re.search(r'Target="([^"]+)"', alvo).group(1).lstrip("/").replace("xl/", "")
    sh = zin.read(caminho).decode("utf-8")
    st = zin.read("xl/styles.xml").decode("utf-8")

    # ---- estilos: duas cores novas e uma cópia de cada estilo usado, com a cor
    nf = int(re.search(r'<fills count="(\d+)"', st).group(1))
    fills = ''.join(f'<fill><patternFill patternType="solid"><fgColor rgb="{c}"/><bgColor indexed="64"/></patternFill></fill>' for c in (COR_VERDE, COR_VERMELHO))
    st = re.sub(r'<fills count="\d+">', f'<fills count="{nf + 2}">', st, count=1).replace("</fills>", fills + "</fills>", 1)
    bloco = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, re.S)
    xfs = re.findall(r'<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>', bloco.group(2), re.S)
    novos, cache = [], {}

    def estilo(s_antigo, cor):
        k = (s_antigo, cor)
        if k not in cache:
            xf = xfs[s_antigo]
            fid = nf + (0 if cor == "verde" else 1)
            cab = re.match(r'<xf\b[^>/]*', xf).group(0)
            cab2 = re.sub(r'\sfillId="\d+"', "", cab); cab2 = re.sub(r'\sapplyFill="\d"', "", cab2)
            xf2 = xf.replace(cab, f'{cab2} fillId="{fid}" applyFill="1"', 1)
            cache[k] = len(xfs) + len(novos)
            novos.append(xf2)
        return cache[k]

    def estilo_de(celula):
        r = re.search(r'\ss="(\d+)"', re.match(r"<c\b[^>/]*", celula).group(0))
        return int(r.group(1)) if r else 0

    def pinta(celula, cor):
        cab = re.match(r"<c\b[^>/]*", celula).group(0)
        s2 = estilo(estilo_de(celula), cor)
        cab2 = re.sub(r'\ss="\d+"', "", cab) + f' s="{s2}"'
        return cab2 + celula[len(cab):]

    def inline(ref, s, texto):
        return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{_xml_texto(texto)}</t></is></c>'

    # ---- linhas existentes
    linhas = {int(mm.group(1)): mm for mm in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', sh, re.S)}
    ult = max(linhas)
    sd_ini = sh.index("<sheetData>")
    saidas_por_linha = {}
    for x in saidas:
        if x.get("linha") and int(x["linha"]) in linhas:
            saidas_por_linha[int(x["linha"])] = "SAIU? " + x["detalhe"] + f" (conferir com a marca {x['Marca']})." + NOTA_SAIDA.get(x["Marca"].upper(), "")
    novo_sheet = sh
    # cabeçalho: coluna de status
    cab = linhas[1].group(0)
    cel_i1 = re.findall(r'<c\b[^>]*?(?:/>|>.*?</c>)', cab, re.S)[-1]
    s_cab = estilo_de(cel_i1)
    cab_novo = re.sub(r'\sspans="[^"]*"', "", cab).replace("</row>", inline("J1", s_cab, COL_STATUS) + inline("L1", estilo(0, "verde"), "Verde = possível loja nova (grupo estimado; conferir)") +
                                                              inline("M1", estilo(0, "vermelho"), "Vermelho = operação que não aparece no site da marca (conferir)") + "</row>")
    novo_sheet = novo_sheet.replace(cab, cab_novo, 1)
    for r_, texto in saidas_por_linha.items():
        orig = linhas[r_].group(0)
        if orig.endswith("/>"):
            continue
        celulas = re.findall(r'<c\b[^>]*?(?:/>|>.*?</c>)', orig, re.S)
        corpo = "".join(pinta(c, "vermelho") for c in celulas)
        s_status = estilo(0, "vermelho")
        abre = re.sub(r'\sspans="[^"]*"', "", re.match(r"<row\b[^>]*>", orig).group(0))
        novo_sheet = novo_sheet.replace(orig, abre + corpo + inline(f"J{r_}", s_status, texto) + "</row>", 1)

    # ---- linhas novas (verde), no final; estilo de cada coluna copiado da última linha de dados
    modelo = re.findall(r'<c\b[^>]*?(?:/>|>.*?</c>)', linhas[ult].group(0), re.S)
    est_col = {re.match(r'<c r="([A-Z]+)', c).group(1): estilo_de(c) for c in modelo}
    def cel(col, r_, valor):
        s_ = estilo(est_col.get(col, 0), "verde")
        if valor == "" or valor is None or (isinstance(valor, float) and valor != valor):
            return f'<c r="{col}{r_}" s="{s_}"/>'
        if isinstance(valor, (int, float)):
            return f'<c r="{col}{r_}" s="{s_}"><v>{valor}</v></c>'
        return inline(f"{col}{r_}", s_, valor)
    novas_linhas, r_ = [], ult
    for x in novas:
        r_ += 1
        grupo = x.get("grupo_provavel") or "(a confirmar)"
        det = ("NOVA - " + x.get("detalhe", "")).strip(" -")
        if x.get("lat") in ("", None):
            det += " | sem coordenada no site (preencher)"
        vals = {"A": grupo, "B": x["Marca"], "C": x.get("cidade", ""), "D": (UF_POR_NOME.get(norm(x.get("uf", "")), x.get("uf", "")) if len(str(x.get("uf", ""))) > 2 else x.get("uf", "")), "E": (x.get("endereco") or x.get("loja_no_site", "")),
                "F": "", "G": "", "H": x.get("lng", ""), "I": x.get("lat", "")}
        corpo = "".join(cel(c, r_, v) for c, v in vals.items()) + inline(f"J{r_}", estilo(est_col.get("A", 0), "verde"), det)
        novas_linhas.append(f'<row r="{r_}">{corpo}</row>')
    novo_sheet = novo_sheet.replace("</sheetData>", "".join(novas_linhas) + "</sheetData>", 1)
    novo_sheet = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:M{max(r_, ult)}"/>', novo_sheet, count=1)
    novo_sheet = novo_sheet.replace("</cols>", '<col min="10" max="10" width="70" customWidth="1"/><col min="12" max="13" width="62" customWidth="1"/></cols>', 1)

    # ---- grava estilos e escreve o novo arquivo (as demais partes são copiadas sem alteração)
    st = st.replace(bloco.group(0), f'<cellXfs count="{len(xfs) + len(novos)}">' + bloco.group(2) + "".join(novos) + "</cellXfs>", 1)
    saida = os.path.join(pasta, _nome_saida(plan))
    tmp = saida + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == caminho:
                dados = novo_sheet.encode("utf-8")
            elif item.filename == "xl/styles.xml":
                dados = st.encode("utf-8")
            zout.writestr(item, dados)
    zin.close()
    os.replace(tmp, saida)
    return saida


def executar(lista=None, marcas=None, sem_download=None, pasta=None):
    """Roda tudo. Devolve {"arquivo","pasta","abas","avisos","resumo"}. Usado pelo terminal e pelo painel."""
    global LOGFILE
    pasta = pasta or f"saida_{datetime.date.today():%Y-%m-%d}"
    os.makedirs(pasta, exist_ok=True)
    LOGFILE = os.path.join(pasta, "log.txt")
    log(f"Início: {datetime.datetime.now():%d/%m/%Y %H:%M:%S}")
    plan = achar_planilha(lista)
    log(f"Planilha da rede: {plan}")
    L = pd.read_excel(plan, sheet_name=ABA_LISTA)
    L.columns = [str(c).strip() for c in L.columns]
    faltam = [c for c in COLUNAS_LISTA if c not in L.columns]
    if faltam:
        raise SystemExit(f"A aba '{ABA_LISTA}' não tem as colunas: {faltam}. Se o nome mudou, ajuste COLUNAS_LISTA no topo do script.")
    L["_linha"] = L.index + 2                      # linha na planilha original (cabeçalho = linha 1)
    if COL_STATUS in L.columns:                    # arquivo já comparado antes: ignora as linhas que o programa tinha acrescentado
        L = L[~L[COL_STATUS].astype(str).str.upper().str.startswith("NOVA")]
    L = L[L["MARCA"].notna()].reset_index(drop=True)
    L["MARCA"] = L["MARCA"].astype(str).str.strip()
    L["LATITUDE"] = L.LATITUDE.map(num); L["LONGITUDE"] = L.LONGITUDE.map(num)
    for c in ("CIDADE", "UF", "ENDEREÇO", "NOME DO GRUPO"):
        L[c] = L[c].fillna("").astype(str).str.strip()
    grupos_todos = set(L["NOME DO GRUPO"].map(norm)) - {""}
    acomp = {m[1].upper() for m in MARCAS}           # só confere a qualidade das marcas que o programa acompanha
    qual = qualidade_lista(L)
    qual = qual[qual.marca.str.upper().isin(acomp)].reset_index(drop=True)
    if len(qual):
        log(f"Atenção: {len(qual)} problema(s) na lista da rede (aba 'Qualidade da lista').")

    hoje_nome = f"{datetime.date.today():%Y-%m-%d}.json"
    hist = carregar_historico(hoje_nome)
    hist_m = hist.get("marcas", {})
    so = {m.strip().upper() for m in marcas} if marcas else None

    resumo, novas, fechadas, divergs, outras, avisos = [], [], [], [], [], []
    pintar_novas, pintar_saidas, pintar_nao = [], [], []   # para a planilha devolvida (marcas com leitura confiável)
    falhas = []                                    # marcas com erro ou que usaram dados antigos (para "tentar de novo só essas")
    snap = {}
    if not sem_download:
        pend = [m for m in estado_manuais() if m["precisa_atencao"] and (not so or m["chave"] in so)]
        if pend:
            log("ARQUIVOS MANUAIS a conferir antes de continuar (ou a marca ficará sem dados/desatualizada):")
            for m in pend:
                txt = f"manual/{m['arquivo']}: {m['situacao']}"
                log(f"  - {m['chave']}: {txt}")
                avisos.append({"Marca": m["chave"], "Tipo": "ATENÇÃO", "Texto": f"Arquivo manual {txt}. Atualize (passo a passo no LEIA-ME)."})
    baixados = {}
    for chave, marca, baixar, comparar in MARCAS:
        if so and chave not in so and marca.upper() not in so:
            continue
        log(f"\n=== {marca} ===")
        lm = L[L.MARCA.str.upper() == marca.upper()]
        linha = {"Marca": marca, "Lojas no site": "", "Variação vs. mês anterior": "", "Operações na lista": len(lm), "Iguais": "",
                 "Endereço divergente": "", "Possíveis novas": "", "Possíveis fechadas": "", "Situação": ""}
        antigo = ""
        try:
            if sem_download:
                arq = csv_mais_recente(sem_download, chave)
                if not arq:
                    raise RuntimeError(f"não achei lojas_{chave}_*.csv em {sem_download}")
                S = pd.read_csv(arq, sep=";", dtype=str).fillna("")
                origem = f"CSV {os.path.basename(arq)}"
            else:
                log("  baixando...")
                cp = os.path.join(CACHE_DIR, f"lojas_{chave}.csv")
                try:
                    S = baixar_com_retentativa(baixar, byd=baixados.get("BYD"))
                    origem = S.attrs.get("origem", "site")
                    S.to_csv(os.path.join(pasta, f"lojas_{chave}_{HOJE:%Y-%m-%d}.csv"), sep=";", index=False, encoding="utf-8-sig")
                    if not S.attrs.get("parcial"):          # só guarda no cache leituras completas
                        os.makedirs(CACHE_DIR, exist_ok=True)
                        S.to_csv(cp, sep=";", index=False, encoding="utf-8-sig")
                except Exception as e_dl:
                    if not os.path.exists(cp):
                        raise
                    S = pd.read_csv(cp, sep=";", dtype=str).fillna("")
                    quando = datetime.datetime.fromtimestamp(os.path.getmtime(cp)).strftime("%d/%m/%Y")
                    motivo = (str(e_dl).splitlines() or [""])[0][:200]
                    antigo = f"usando dados ANTIGOS de {quando}, porque o download falhou ({motivo})"
                    origem = f"CACHE de {quando}"
                    log(f"  AVISO: {antigo}")
            baixados[chave] = S
            log(f"  {len(S)} registros ({origem})")
            if len(S) == 0:
                raise ErroAcesso("mudou", f"o site de {marca} respondeu, mas sem nenhuma loja (o formato pode ter mudado).")
            if len(lm) == 0:
                linha["Situação"] = "marca sem operações na lista; só baixou"
                resumo.append(linha); continue
            res = comparar(S, lm, grupos_todos=grupos_todos)
            alertas = []
            if antigo:
                alertas.append(antigo)
            parcial = S.attrs.get("parcial")
            if parcial:
                alertas.append(f"leitura parcial: {parcial}")
            if origem.startswith("arquivo manual/"):
                d = idade_manual(origem.split("/", 1)[1])
                if d is not None and d > IDADE_MAX_MANUAL:
                    alertas.append(f"o arquivo manual tem {d} dias; os dados podem estar desatualizados")
            if antigo:
                falhas.append(chave)
            esp = ESPERADO.get(chave)
            ant = (hist_m.get(chave) or {}).get("n_site")
            base = ant or esp                       # referência: leitura do mês anterior; se não houver, o valor esperado
            incompleta = (bool(base) and res["n_site"] < 0.6 * base) or bool(parcial)
            if incompleta:
                alertas.append(f"leitura possivelmente INCOMPLETA ({res['n_site']} lojas; referência ~{base}). As 'possíveis fechadas' foram ocultadas porque não são confiáveis" if not parcial else "as 'possíveis fechadas' foram ocultadas porque a leitura foi parcial")
                avisos.append({"Marca": marca, "Tipo": "ATENÇÃO", "Texto": f"Leitura incompleta: {res['n_site']} lojas contra ~{base} de referência. Rode de novo; se persistir, o site pode ter mudado (use o plano B do LEIA-ME)."})
                res["fechadas"] = []
            elif base and not (0.7 * base <= res["n_site"] <= 1.3 * base):
                alertas.append(f"{res['n_site']} lojas no site contra ~{base} de referência")
                avisos.append({"Marca": marca, "Tipo": "ATENÇÃO", "Texto": f"O site mostrou {res['n_site']} lojas e a referência é ~{base}; confira se o site mudou ou se a coleta veio incompleta"})
            if ant:
                linha["Variação vs. mês anterior"] = f"{res['n_site'] - ant:+d}"
            if not incompleta and res["iguais"] + len(res["diverg"]) + len(res["fechadas"]) != res["n_lista"]:
                alertas.append("a conta não fecha (iguais + divergentes + fechadas ≠ operações da lista)")
                avisos.append({"Marca": marca, "Tipo": "ATENÇÃO", "Texto": f"Conferência: iguais {res['iguais']} + divergentes {len(res['diverg'])} + fechadas {len(res['fechadas'])} ≠ {res['n_lista']} operações da lista"})
            if antigo:                              # dados antigos (cache): não pinta nada da marca
                pintar_nao.append(marca)
            elif not incompleta:
                pintar_novas += [{"Marca": marca, **x} for x in res["novas"]]
                pintar_saidas += [{"Marca": marca, **x} for x in res["fechadas"]]
            else:
                pintar_nao.append(marca)
            for k, dest in (("novas", novas), ("fechadas", fechadas), ("diverg", divergs), ("outras", outras)):
                for x in res[k]:
                    dest.append({"Marca": marca, **x})
            linha.update({"Lojas no site": res["n_site"], "Operações na lista": res["n_lista"], "Iguais": res["iguais"], "Endereço divergente": len(res["diverg"]),
                          "Possíveis novas": len(res["novas"]), "Possíveis fechadas": "oculto" if incompleta else len(res["fechadas"]),
                          "Situação": ("ATENÇÃO: " + "; ".join(alertas)) if alertas else f"ok ({origem})"})
            for n in res["notas"]:
                avisos.append({"Marca": marca, "Tipo": "observação", "Texto": n})
            if not antigo:                          # dados antigos (cache) não entram no histórico
              snap[chave] = {"n_site": res["n_site"], "itens": sorted(set(
                [chave_item("Possível loja nova", marca, x["cidade"], x["grupo_provavel"]) for x in res["novas"]] +
                [chave_item("Possível fechada", marca, x["cidade"], x["grupo"]) for x in res["fechadas"]] +
                [chave_item("Endereço divergente", marca, x["cidade"], x["grupo"]) for x in res["diverg"]]))}
            log(f"  iguais {res['iguais']} | divergentes {len(res['diverg'])} | novas {len(res['novas'])} | fechadas {len(res['fechadas'])}"
                + ("  [ATENÇÃO: " + "; ".join(alertas) + "]" if alertas else ""))
        except Exception as e:
            msg = str(e).splitlines()[0][:400] if str(e) else e.__class__.__name__
            if isinstance(e, (KeyError, IndexError, TypeError, AttributeError, json.JSONDecodeError)):
                msg = f"o formato dos dados de {marca} não é o esperado ({e.__class__.__name__}: {str(e)[:80]}). O site provavelmente mudou; avise quem mantém o programa (veja log.txt)."
            falhas.append(chave)
            pintar_nao.append(marca)
            linha["Situação"] = "ERRO: " + msg
            avisos.append({"Marca": marca, "Tipo": "ERRO", "Texto": msg})
            log(f"  ERRO: {msg}")
            if LOGFILE:
                try:
                    with open(LOGFILE, "a", encoding="utf-8") as f:
                        f.write(traceback.format_exc() + "\n")     # detalhe técnico só no arquivo de log
                except Exception:
                    pass
            if os.environ.get("DEBUG"):
                traceback.print_exc()
        resumo.append(linha)

    ja = {m[1].upper() for m in MARCAS}
    for marca, n in L[~L.MARCA.str.upper().isin(ja)].groupby("MARCA").size().items():
        if marca.upper() not in SEM_COLETA:      # marcas que não acompanhamos não aparecem
            continue
        motivo = SEM_COLETA[marca.upper()]
        resumo.append({"Marca": marca, "Lojas no site": "", "Variação vs. mês anterior": "", "Operações na lista": int(n), "Iguais": "", "Endereço divergente": "",
                       "Possíveis novas": "", "Possíveis fechadas": "", "Situação": "não automatizada: " + motivo})

    # --- "Para confirmar": só o que a equipe de rede precisa decidir, com marca de NOVO / já aparecia
    def status(tipo, marca, cidade, grupo):
        if not hist_m:
            return "sem histórico"
        return "já aparecia antes" if chave_item(tipo, marca, cidade, grupo) in set((hist_m.get(chave_de(marca)) or {}).get("itens", [])) else "NOVO neste mês"
    chave_de = lambda m: next((c for c, nm, _, _ in MARCAS if nm == m), m)
    conf = []
    for x in novas:
        conf.append({"Marca": x["Marca"], "Tipo": "Possível loja nova", "Grupo": x["grupo_provavel"], "Cidade": x["cidade"],
                     "Detalhe": x["loja_no_site"] + (" · " + x["endereco"] if x["endereco"] else ""), "Status": status("Possível loja nova", x["Marca"], x["cidade"], x["grupo_provavel"])})
    for x in fechadas:
        conf.append({"Marca": x["Marca"], "Tipo": "Possível fechada", "Grupo": x["grupo"], "Cidade": x["cidade"],
                     "Detalhe": f'{x["endereco_na_lista"]} · {x["detalhe"]}', "Status": status("Possível fechada", x["Marca"], x["cidade"], x["grupo"])})
    for x in divergs:
        conf.append({"Marca": x["Marca"], "Tipo": "Endereço divergente", "Grupo": x["grupo"], "Cidade": x["cidade"],
                     "Detalhe": f'lista: {x["endereco_lista"]} · site: {x["loja_no_site"]} {x["endereco_no_site"]}', "Status": status("Endereço divergente", x["Marca"], x["cidade"], x["grupo"])})
    colconf = ["Marca", "Tipo", "Grupo", "Cidade", "Detalhe", "Status"]
    dconf = pd.DataFrame(conf, columns=colconf)
    dconf["Confirmado? (sim/não)"] = ""; dconf["Observação"] = ""

    arq = os.path.join(pasta, f"Comparacao_Rede_{HOJE:%Y-%m-%d}.xlsx")
    vazio = lambda c: pd.DataFrame(columns=c)
    abas = {
        "Resumo geral": pd.DataFrame(resumo),
        "Para confirmar": dconf,
        "Possiveis lojas novas": pd.DataFrame(novas) if novas else vazio(["Marca", "loja_no_site", "cidade", "uf", "endereco", "grupo_provavel", "detalhe"]),
        "Possiveis fechadas": pd.DataFrame(fechadas) if fechadas else vazio(["Marca", "grupo", "cidade", "uf", "endereco_na_lista", "detalhe"]),
        "Endereco divergente": pd.DataFrame(divergs) if divergs else vazio(["Marca", "grupo", "cidade", "endereco_lista", "loja_no_site", "endereco_no_site", "km", "detalhe"]),
        "Outros grupos": pd.DataFrame(outras) if outras else vazio(["Marca", "loja_no_site", "cidade", "uf", "endereco", "grupo_provavel", "detalhe"]),
        "Qualidade da lista": qual,
        "Avisos": pd.DataFrame(avisos) if avisos else vazio(["Marca", "Tipo", "Texto"]),
    }
    arq = salvar_xlsx(arq, abas)

    lista_comp = None
    try:
        lista_comp = gerar_lista_comparada(plan, pasta, pintar_novas, pintar_saidas)
        log(f"\nPlanilha para a equipe de rede (com cores): {lista_comp}")
        if pintar_nao:
            log("  Sem destaque nestas marcas (erro, leitura incompleta ou dados antigos): " + ", ".join(dict.fromkeys(pintar_nao)))
    except Exception as e:
        log(f"\nATENÇÃO: não consegui gerar a planilha com cores ({e}). O relatório de comparação foi gerado normalmente.")
        avisos.append({"Marca": "(geral)", "Tipo": "ERRO", "Texto": f"planilha com cores não gerada: {e}"})
        with open(LOGFILE, "a", encoding="utf-8") as f:
            f.write(traceback.format_exc() + "\n")
    if pintar_nao:
        avisos.append({"Marca": ", ".join(dict.fromkeys(pintar_nao)), "Tipo": "ATENÇÃO", "Texto": "Sem destaque na planilha com cores (erro, leitura incompleta ou dados antigos). Resolva e rode de novo só estas marcas."})

    if not sem_download and snap:                  # guarda o resultado de hoje para comparar no mês que vem
        try:
            os.makedirs(HIST_DIR, exist_ok=True)
            marcas_h = dict(hist_m); marcas_h.update(snap)
            with open(os.path.join(HIST_DIR, hoje_nome), "w", encoding="utf-8") as f:
                json.dump({"data": str(datetime.date.today()), "marcas": marcas_h}, f, ensure_ascii=False)
        except Exception as e:
            log(f"  (não consegui gravar o histórico: {e})")

    log("\n" + "=" * 70)
    log(abas["Resumo geral"].to_string(index=False))
    log("=" * 70)
    log(f"\nPronto! Planilha: {arq}")
    log(f"CSVs de cada marca e log.txt: pasta {pasta}")
    erros = [x for x in avisos if x["Tipo"] == "ERRO"]
    if erros:
        log(f"\nATENÇÃO: {len(erros)} marca(s) com erro (veja a aba Avisos): " + ", ".join(x["Marca"] for x in erros))
    if falhas:
        log("\nMarcas com problema neste mês: " + ", ".join(dict.fromkeys(falhas)) + ". Depois de corrigir (ex.: salvar o arquivo manual), rode só elas: --marcas " + ",".join(dict.fromkeys(falhas)))
    return {"arquivo": arq, "pasta": pasta, "abas": abas, "avisos": avisos, "resumo": resumo, "lista_comparada": lista_comp, "pintar_nao": list(dict.fromkeys(pintar_nao)), "falhas": list(dict.fromkeys(falhas))}


def main():
    ap = argparse.ArgumentParser(description="Atualiza a comparação da rede de concessionárias (todas as marcas).")
    ap.add_argument("--lista", help="planilha da rede (padrão: procura na pasta)")
    ap.add_argument("--marcas", help="só estas marcas, separadas por vírgula (ex.: BYD,GWM,MG)")
    ap.add_argument("--sem-download", metavar="PASTA", help="não baixa; usa os CSVs lojas_<MARCA>_*.csv desta pasta")
    ap.add_argument("--diagnostico", action="store_true", help="só testa a conexão com os sites das marcas")
    a = ap.parse_args()
    if a.diagnostico:
        diagnostico_conexao(); return
    executar(a.lista, a.marcas.split(",") if a.marcas else None, a.sem_download)


if __name__ == "__main__":
    main()