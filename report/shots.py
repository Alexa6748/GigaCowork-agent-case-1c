import openpyxl, html, sys
from datetime import date, datetime
from playwright.sync_api import sync_playwright
src=sys.argv[1]
wb=openpyxl.load_workbook(src)
def tbl(name, cols=None, rows=None, maxlen=None):
    ws=wb[name]; out=[]
    for r_i,row in enumerate(ws.iter_rows()):
        if rows and r_i>rows: break
        cells=[]
        for c_i,c in enumerate(row):
            if cols and c_i not in cols: continue
            v=c.value
            if isinstance(v,(date,datetime)): v=v.strftime('%d.%m.%Y')
            v='' if v is None else str(v)
            if maxlen and len(v)>maxlen.get(c_i,999): v=v[:maxlen[c_i]]+'…'
            fill=c.fill.fgColor.rgb if c.fill and c.fill.fill_type=='solid' else None
            st=''
            if r_i==0: st='background:#1F4E78;color:#fff;font-weight:600'
            elif fill and fill not in ('00000000',): st='background:#'+fill[-6:]
            cells.append('<td style="%s">%s</td>'%(st,html.escape(v)))
        out.append('<tr><th>%d</th>%s</tr>'%(r_i+1,''.join(cells)))
    return '<table>%s</table>'%''.join(out)
css='body{margin:0;font:12px Arial;background:#fff}.tab{background:#e8e8e8;padding:4px 8px;font:12px Arial;color:#333;border-bottom:1px solid #bbb}.tab b{background:#fff;padding:3px 10px;border:1px solid #bbb;border-bottom:none}table{border-collapse:collapse}td,th{border:1px solid #c8c8c8;padding:3px 6px;vertical-align:top;max-width:520px}th{background:#f0f0f0;color:#666;font-weight:normal}'
jobs=[('Новые строки Cloud',None,19,None,'new'),('Журнал решений',[1,3,4,8,9,11,13],12,{13:260},'journal'),('Требует внимания',[0,3,5,6,7],9,None,'att'),('Повторные отправки',None,5,None,'rep')]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(device_scale_factor=2, viewport={'width':1400,'height':400})
    for name,cols,rows,ml,fn in jobs:
        tabs=''.join(('<b>%s</b>' if s==name else ' %s ')%s for s in wb.sheetnames)
        pg.set_content('<html><head><style>%s</style></head><body><div class="tab">%s</div>%s</body></html>'%(css,tabs,tbl(name,cols,rows,ml)))
        el=pg.locator('body'); el.screenshot(path='img/%s.png'%fn)
    b.close()
print('ok')
