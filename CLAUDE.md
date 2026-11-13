# Rede de concessionárias — guia para quem mantém este projeto

Projeto mensal da BMW e-mobility: compara a lista de operações da equipe de rede (planilha, aba "Lista Operações") com as lojas
dos sites oficiais das marcas chinesas. Quem **roda** o mês usa só o painel (`Iniciar_Rede.bat`) e o `LEIA-ME_rede.txt`.
Este arquivo é para quem **altera o código** (humano ou assistente de IA). Fale com o usuário em português.

## Arquivos
- `atualizar_rede.py` — todo o programa (downloads, comparações, planilhas). Roda sozinho no terminal: `python atualizar_rede.py`.
- `app_rede.py` — painel local (navegador, 127.0.0.1) que chama `atualizar_rede.executar(...)`. Sem lógica de negócio.
- `marcas_extras.csv` — cadastro de marcas com API JSON simples (sem programar).
- `manual/` — arquivos que uma pessoa salva quando o site bloqueia (`gac.json`, `zeekr.json` sempre; os demais só se o site falhar).
- `cache/` — última leitura completa de cada marca (reserva se o download falhar). `historico/` — resumo de cada mês. Não apagar.
- `saida/AAAA-MM-DD_HHMM/` — resultados de cada rodada (+ `log.txt`).

## Fluxo de `executar()`
1. Lê a lista (`ABA_LISTA`, colunas em `COLUNAS_LISTA`); guarda a linha original em `_linha`; ignora linhas `NOVA` de arquivo já comparado.
2. Para cada marca em `MARCAS` (chave, nome na lista, função de download, função de comparação): baixa (com retentativa) → se falhar usa
   `cache/` marcado como "dados antigos" → compara → confere (leitura incompleta < 60% do mês anterior/`ESPERADO`, variação ±30%, conta fecha).
3. Gera `Comparacao_Rede_*.xlsx` (conferência) e `REP_AAMMDD_*.xlsx` (cópia da lista com verde = novas, vermelho = saídas, coluna STATUS).
4. Grava `historico/` e `cache/` só com leituras completas e não-antigas.

## Planilha com cores (`gerar_lista_comparada`)
Edita **direto o XML** da aba da lista dentro do .xlsx (só `sheet` da lista e `styles.xml` mudam). NÃO trocar por openpyxl/pandas:
o arquivo original tem imagens, rich data, links externos e formatação condicional que o openpyxl apagaria. Cuidados já tratados:
células `<c .../>` auto-fechadas, ordem das colunas na linha, contagem de fills/cellXfs. Depois de mexer, validar: abrir com LibreOffice/openpyxl,
checar que só 2 partes do zip mudaram e que as linhas/células seguem em ordem. A aba "Qntd Operações" usa COUNTIFS na lista inteira
(as linhas verdes entram na contagem quando o Excel recalcula) — avisar a equipe de rede.

## Como adicionar uma marca
1. Marca na lista da rede com o nome exato em `nome_na_lista`.
2. API JSON simples → uma linha em `marcas_extras.csv` (url, `caminho_lista` tipo `data.list`, campos name/cidade/uf/address/lat/lng, `metodo` =
   `distancia` (usa lat/lng, raio `km`) ou `nome_cidade` (sem coordenadas)). Linhas inválidas são ignoradas com aviso no log.
3. Site HTML/assinado → escrever `baixar_<marca>()` em `atualizar_rede.py` que devolva DataFrame com colunas `name, cidade, uf, address, lat, lng, vende`
   (`vende` = "sim"/"nao (pos-venda)"), pôr `df.attrs["origem"]`, e registrar em `MARCAS` com `_dist(...)` ou `_nc(...)`. Usar `pagina(url, "arquivo.html")`
   para ter o plano B manual automático. Cadastrar `ESPERADO` e `APELIDOS`. Se precisar de arquivo manual, listar em `MANUAIS_OBRIGATORIOS`/`MANUAIS_RESERVA`
   e em `MANUAIS` do `app_rede.py`.
4. Testar com `python atualizar_rede.py --marcas CHAVE` e com um CSV (`--sem-download`).
Para outra área (ex.: Motos): precisa de outra lista de operações (mesmo formato) — a ideia é uma coluna "área" no cadastro e uma lista por área.

## Acesso aos sites (regras firmes)
- `http()` classifica falhas em português (`ErroAcesso`) e só repete as passageiras (rede, 5xx, 429). 401/403/404/certificado/proxy não repetem.
- **Nunca** forjar assinatura, burlar CAPTCHA/proteção de robô, nem pedir cookies/tokens ao usuário. GAC e Zeekr exigem assinatura do navegador:
  o caminho é o arquivo manual (`manual/*.json`, resposta copiada do F12). Não imprimir chaves/tokens em log. Não guardar contatos (nomes/e-mails) da Zeekr em saídas.
- `--diagnostico` / botão "Testar conexão" testa cada site sem baixar tudo.

## Armadilhas conhecidas
- Zeekr/GAC: latitude e longitude vêm trocadas em algumas lojas (corrigido quando lat < -34.5).
- GWM: só lojas com ponto próprio (filtro `PADRAO_SHOPPING`, heurístico); o site devolve lojas por CEP (varredura em largura, com disjuntor após 8 falhas seguidas).
- Jetour: o site mostra ~60 de ~69 lojas — "saídas" da Jetour precisam de confirmação.
- Grupo de cada loja é estimativa (nome/e-mail/CNPJ); só o CNPJ da GWM é confiável. Denza/Zeekr: grupo "(a confirmar)".
- Lista da rede tem erros de coordenada (aba "Qualidade da lista"); isso gera falsos "saiu".
