"""Lê a planilha 'Calendário de Ações' do SESMT e gera dados/acoes.json para o painel."""
import openpyxl, json, sys, os, warnings, datetime
warnings.filterwarnings('ignore')
ORIG={'CIPA':'CIPA','AUD. COMPORT..':'AUD','COMITÊ SEG.':'COMITE','CAMPANHA':'CAMPANHA','ACIDENTE':'ACIDENTE'}
DEP={'ELETRICA':'MANUTENÇÃO ELÉTRICA','MANUTENÇÃO':'MANUTENÇÃO MEC.'}
d=lambda v: v.strftime('%Y-%m-%d') if hasattr(v,'strftime') else None
txt=lambda v: ' '.join(str(v or '').split())

def linha(n,orig,do,desc,local,resp,prev,term,obs,hist):
    o=ORIG.get(txt(orig))
    if not o:
        print('  aviso: origem desconhecida', n, orig); return None
    dep=txt(resp) or 'SEM RESPONSÁVEL'
    return dict(n=n,o=o,do=d(do),desc=txt(desc),local=txt(local),dep=DEP.get(dep,dep),
                prev=d(prev),term=d(term),obs=txt(obs),hist=hist)

import re
FOTO_RE = re.compile(r'^(H?\d+)(?:[-_ ].*)?\.(jpe?g|png|webp)$', re.I)
def fotos_por_acao(pasta):
    """fotos/274.jpg, fotos/274-2.jpg, fotos/H12.png -> {'274': ['fotos/274.jpg', ...]}"""
    m = {}
    if pasta and os.path.isdir(pasta):
        for f in sorted(os.listdir(pasta)):
            r = FOTO_RE.match(f)
            if r: m.setdefault(r.group(1).upper(), []).append(f'fotos/{f}')
    return m

def carregar(caminho):
    try:
        return json.load(open(caminho, encoding='utf-8'))
    except Exception:
        return None

DIAS_GUARDAR = 45
def registrar_novidades(anterior, atuais, caminho, quando):
    """Compara a versão publicada com a nova e registra o que mudou em dados/novidades.json."""
    nov = carregar(caminho) or {'eventos': []}
    ev = []
    if anterior and anterior.get('acoes'):
        ant = {str(a['n']): a for a in anterior['acoes'] if not a.get('hist')}
        nov_ = {str(a['n']): a for a in atuais if not a.get('hist')}
        hist_desc = {a['desc']: a for a in atuais if a.get('hist')}
        base = lambda a: dict(n=a['n'], desc=a['desc'], o=a['o'], local=a['local'], dep=a['dep'])
        for k, a in nov_.items():
            b = ant.get(k)
            if b is None:
                ev.append(dict(base(a), tipo='nova', para=a.get('prev')))
            elif a.get('term') and not b.get('term'):
                ev.append(dict(base(a), tipo='concluida', para=a['term']))
            elif not a.get('term') and b.get('term'):
                ev.append(dict(base(a), tipo='reaberta', de=b['term']))
            elif a.get('prev') != b.get('prev'):
                ev.append(dict(base(a), tipo='data', de=b.get('prev'), para=a.get('prev')))
        for k, b in ant.items():
            if k in nov_: continue
            h = hist_desc.get(b['desc'])
            if h and not b.get('term'):
                ev.append(dict(base(b), tipo='concluida', para=h.get('term')))
            elif not h:
                ev.append(dict(base(b), tipo='removida'))
    for e in ev: e['quando'] = quando
    limite = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=DIAS_GUARDAR)).isoformat()
    eventos = [e for e in nov.get('eventos', []) if e.get('quando', '') >= limite] + ev
    json.dump({'atualizado_em': quando if ev else nov.get('atualizado_em'), 'eventos': eventos},
              open(caminho, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f'{len(ev)} novidade(s) registrada(s)')

def main(xlsx, saida, pasta_fotos='fotos'):
    wb=openpyxl.load_workbook(xlsx,data_only=True)
    out=[]
    # aba principal: MASTER PLAN
    master = next((w for w in wb.worksheets if w.title.upper().startswith('MASTER PLAN')), wb.worksheets[0])
    concl = next((w for w in wb.worksheets if w.title.upper().startswith('CONCLU')), None)
    for r in master.iter_rows(min_row=6,values_only=True):
        if not r[4]: continue
        a=linha(r[1],r[2],r[3],r[4],r[7],r[8],r[9],r[11],r[14],False)
        if a: out.append(a)
    # aba CONCLUÍDOS (histórico) — ignora o que já está na aba principal
    ja={a['desc'] for a in out}
    nh=0
    anterior = carregar(saida)
    if concl is not None:
        for r in concl.iter_rows(min_row=6,values_only=True):
            if not r[4] or txt(r[4]) in ja: continue
            a=linha(f'H{r[1]}',r[2],r[3],r[4],r[7],r[8],r[9],r[10],r[12],True)
            if a: out.append(a); nh+=1
    elif anterior:
        # planilha sem a aba CONCLUÍDOS (ex.: versão do SharePoint): mantém o histórico já publicado
        for a in anterior.get('acoes', []):
            if a.get('hist') and a['desc'] not in ja:
                a = {k: v for k, v in a.items() if k != 'fotos'}; out.append(a); nh += 1
        print('  aba CONCLUÍDOS ausente: histórico anterior mantido')
    fm = fotos_por_acao(pasta_fotos)
    for a in out:
        a['fotos'] = fm.get(str(a['n']).upper(), [])
    print(f"{sum(1 for a in out if a['fotos'])} ações com foto")
    pacote={'gerado_em':datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds'),
            'arquivo':os.path.basename(xlsx),'acoes':out}
    os.makedirs(os.path.dirname(saida) or '.',exist_ok=True)
    registrar_novidades(anterior, out, os.path.join(os.path.dirname(saida) or '.', 'novidades.json'), pacote['gerado_em'])
    json.dump(pacote,open(saida,'w',encoding='utf-8'),ensure_ascii=False)
    print(f'{len(out)} ações ({nh} do histórico CONCLUÍDOS)')

if __name__=='__main__':
    import argparse
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--file',required=True,help='planilha .xlsx/.xlsm do Calendário de Ações')
    p.add_argument('--out',default='dados/acoes.json')
    p.add_argument('--fotos',default='fotos',help='pasta com fotos nomeadas pelo nº da ação')
    a=p.parse_args()
    main(a.file,a.out,a.fotos)
