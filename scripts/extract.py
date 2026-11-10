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

def main(xlsx, saida):
    wb=openpyxl.load_workbook(xlsx,data_only=True)
    out=[]
    # aba principal: MASTER PLAN
    for r in wb.worksheets[0].iter_rows(min_row=6,values_only=True):
        if not r[4]: continue
        a=linha(r[1],r[2],r[3],r[4],r[7],r[8],r[9],r[11],r[14],False)
        if a: out.append(a)
    # aba CONCLUÍDOS (histórico) — ignora o que já está na aba principal
    ja={a['desc'] for a in out}
    nh=0
    if len(wb.worksheets)>2:
        for r in wb.worksheets[2].iter_rows(min_row=6,values_only=True):
            if not r[4] or txt(r[4]) in ja: continue
            a=linha(f'H{r[1]}',r[2],r[3],r[4],r[7],r[8],r[9],r[10],r[12],True)
            if a: out.append(a); nh+=1
    pacote={'gerado_em':datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds'),
            'arquivo':os.path.basename(xlsx),'acoes':out}
    os.makedirs(os.path.dirname(saida) or '.',exist_ok=True)
    json.dump(pacote,open(saida,'w',encoding='utf-8'),ensure_ascii=False)
    print(f'{len(out)} ações ({nh} do histórico CONCLUÍDOS)')

if __name__=='__main__':
    import argparse
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--file',required=True,help='planilha .xlsx/.xlsm do Calendário de Ações')
    p.add_argument('--out',default='dados/acoes.json')
    a=p.parse_args()
    main(a.file,a.out)
