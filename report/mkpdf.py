from playwright.sync_api import sync_playwright
import pathlib
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page()
    pg.goto(pathlib.Path('report.html').resolve().as_uri()); pg.wait_for_timeout(500)
    pg.pdf(path='Отчёт_агент_распределения_ФУЛ.pdf',format='A4',print_background=True,display_header_footer=True,header_template='<span></span>',
           footer_template='<div style="font-size:8px;color:#888;width:100%;text-align:center"><span class="pageNumber"></span> / <span class="totalPages"></span></div>',
           margin={'top':'14mm','bottom':'16mm','left':'14mm','right':'14mm'})
    b.close()
