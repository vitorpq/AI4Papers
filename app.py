import os, re, time, csv, shutil
import PySimpleGUI as sg
import pandas as pd
import requests
from urllib.parse import urlparse, unquote
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ======= HTTP (requests) helpers =======

DEFAULT_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "application/pdf,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}

def requests_session():
    s = requests.Session()
    retry = Retry(
        total=5, connect=5, read=5,
        backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["HEAD","GET","OPTIONS"])
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=50, pool_maxsize=50)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update(DEFAULT_HEADERS.copy())
    return s

def ieee_fix_url(url: str) -> str:
    """Converte URLs 'ielx7/.../NNNNNNNN.pdf' para o endpoint oficial stampPDF."""
    if "ieeexplore.ieee.org" not in url:
        return url
    m = re.search(r'/(\d{7,9})\.pdf$', url) or re.search(r'(\d{7,9})', url)
    if m:
        ar = m.group(1).lstrip("0")
        if ar:
            return f"https://ieeexplore.ieee.org/stampPDF/getPDF.jsp?tp=&arnumber={ar}"
    return url

def filename_from_headers(url, resp, default="arquivo.pdf"):
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^\";]+)"?', cd)
    if m:
        name = unquote(m.group(1))
        return name if name.lower().endswith(".pdf") else name + ".pdf"
    path = urlparse(resp.url or url).path
    name = unquote(os.path.basename(path) or default)
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name

def ensure_pdf_response(resp) -> bool:
    ctype = (resp.headers.get("Content-Type") or "").lower()
    return ("pdf" in ctype) or (resp.url.lower().endswith(".pdf"))

def domain_headers(url: str) -> dict:
    dom = urlparse(url).netloc.lower()
    scheme = urlparse(url).scheme or "https"
    origin = f"{scheme}://{dom}"
    h = {"Referer": origin}
    if "mdpi.com" in dom:
        h["Referer"] = "https://www.mdpi.com/"
        h["Accept"] = "application/pdf,*/*;q=0.8"
    elif "jamanetwork.com" in dom:
        h["Referer"] = "https://jamanetwork.com/"
        h["Accept"] = "application/pdf,*/*;q=0.8"
    elif "ieeexplore.ieee.org" in dom:
        h["Referer"] = "https://ieeexplore.ieee.org/"
        h["Accept"] = "application/pdf,*/*;q=0.8"
    return h

def baixar_pdf_requests(sess, url, pasta_destino, throttle_s=0.7):
    """Tenta baixar via requests; lança exceção se não conseguir."""
    os.makedirs(pasta_destino, exist_ok=True)
    fixed = ieee_fix_url(url)
    headers = domain_headers(fixed)

    # preflight
    try:
        sess.head(fixed, headers=headers, allow_redirects=True, timeout=20)
    except requests.RequestException:
        pass

    resp = sess.get(fixed, headers=headers, stream=True, allow_redirects=True, timeout=60)

    # tentativa de “isca de cookie” quando o servidor bloqueia pdf direto
    if resp.status_code in (401, 403):
        base = fixed.split("?")[0]
        if base.endswith("/pdf"):
            page_url = base[:-4]
            try:
                sess.get(page_url, headers=headers, timeout=20)
                resp = sess.get(fixed, headers=headers, stream=True, allow_redirects=True, timeout=60)
            except requests.RequestException:
                pass

    # alguns portais retornam 418 em anti-bot
    if resp.status_code == 418:
        resp.raise_for_status()  # deixamos a exceção subir para acionar fallback

    resp.raise_for_status()
    if not ensure_pdf_response(resp):
        raise requests.HTTPError(f"Conteúdo não parece PDF (Content-Type: {resp.headers.get('Content-Type')})", response=resp)

    nome = filename_from_headers(fixed, resp)
    destino = os.path.join(pasta_destino, nome)
    with open(destino, "wb") as f:
        for chunk in resp.iter_content(1024 * 64):
            if chunk:
                f.write(chunk)

    size = int(resp.headers.get("Content-Length") or 0)
    time.sleep(throttle_s)
    return {"nome": nome, "caminho": destino, "metodo": "requests", "status": resp.status_code, "bytes": size}


# ======= Playwright fallback =======

# Import tardio (lazy) para não exigir Playwright em ambientes sem ele:
def baixar_via_playwright(url, pasta_destino, headless=True, timeout_ms=90000):
    """
    Usa Playwright para navegar como navegador:
    1) goto(url)
    2) se houver evento de download, salva
    3) senão, se a resposta principal for application/pdf, salva o body
    """
    from playwright.sync_api import sync_playwright
    from urllib.parse import urlparse
    os.makedirs(pasta_destino, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        got_download = None
        def _on_download(download):
            nonlocal got_download
            got_download = download
        page.on("download", _on_download)

        # Dica: uma visita rápida à raiz do domínio pode setar cookies úteis
        try:
            parsed = urlparse(url)
            origin = f"{parsed.scheme}://{parsed.netloc}"
            page.goto(origin, timeout=15000)
        except Exception:
            pass

        # IEEE: já ajustamos a URL em 'requests'; aqui só navegamos
        resp = page.goto(url, timeout=timeout_ms, wait_until="domcontentloaded")

        # Caso o site obrigue clique no link do PDF:
        try:
            # tenta clicar no primeiro link que contenha ".pdf"
            page.click("a[href$='.pdf'], a[href*='.pdf?']", timeout=3000)
            # espera um pouco por um possível evento de download
            for _ in range(30):
                if got_download:
                    break
                time.sleep(0.2)
        except Exception:
            pass

        if got_download:
            sug = got_download.suggested_filename or "arquivo.pdf"
            destino = os.path.join(pasta_destino, sug)
            got_download.save_as(destino)
            context.close(); browser.close()
            return {"nome": sug, "caminho": destino, "metodo": "playwright-download", "status": 200, "bytes": os.path.getsize(destino)}

        # Se não houve download, tentamos salvar a resposta principal (inline PDF)
        if resp:
            ct = (resp.headers.get("content-type") or "").lower()
            if "pdf" in ct or page.url.lower().endswith(".pdf"):
                nome = os.path.basename(urlparse(page.url).path) or "arquivo.pdf"
                if not nome.lower().endswith(".pdf"):
                    nome += ".pdf"
                destino = os.path.join(pasta_destino, nome)
                body = resp.body()
                with open(destino, "wb") as f:
                    f.write(body)
                context.close(); browser.close()
                return {"nome": nome, "caminho": destino, "metodo": "playwright-body", "status": resp.status, "bytes": os.path.getsize(destino)}

        context.close(); browser.close()
        raise RuntimeError("Playwright não conseguiu capturar o PDF nesta URL.")

# ======= GUI =======

sg.theme("SystemDefaultForReal")

layout = [
    [sg.Text("Arquivo Excel:"), sg.Input(key="-EXCEL-"), sg.FileBrowse(file_types=(("Excel", "*.xlsx;*.xls"),))],
    [sg.Text("Nome da coluna com URLs:"), sg.Input("url", key="-COL-", size=(20,1))],
    [sg.Text("Pasta de saída:"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
    [sg.Checkbox("Criar ZIP ao final", default=True, key="-ZIP-")],
    [sg.Checkbox("Usar Playwright como fallback", default=True, key="-PWFB-"),
     sg.Checkbox("Headless (sem janela)", default=True, key="-HEADLESS-")],
    [sg.Button("Baixar PDFs", size=(15,1)), sg.Button("Sair")],
    [sg.ProgressBar(max_value=100, orientation="h", size=(40,20), key="-PB-")],
    [sg.Multiline(size=(95,18), key="-LOG-", autoscroll=True, disabled=True)]
]

window = sg.Window("PDF Downloader (Pro)", layout)

def log(window, msg):
    window["-LOG-"].update(msg + "\n", append=True)

def zip_dir(dir_path, zip_path):
    if os.path.exists(zip_path):
        os.remove(zip_path)
    base_name = zip_path[:-4] if zip_path.lower().endswith(".zip") else zip_path
    shutil.make_archive(base_name, "zip", dir_path)
    return base_name + ".zip"

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Sair"):
        break

    if event == "Baixar PDFs":
        excel = values["-EXCEL-"]
        coluna = values["-COL-"].strip()
        outdir = values["-OUT-"]
        use_pw = values["-PWFB-"]
        headless = values["-HEADLESS-"]

        if not excel or not os.path.exists(excel):
            sg.popup_error("Selecione um arquivo Excel válido.")
            continue
        if not coluna:
            sg.popup_error("Informe o nome da coluna que contém as URLs.")
            continue
        if not outdir:
            sg.popup_error("Selecione a pasta de saída.")
            continue

        window["-LOG-"].update("")
        log(window, f"Iniciando… Excel={excel} | coluna={coluna} | saída={outdir}")
        try:
            df = pd.read_excel(excel)
        except Exception as e:
            sg.popup_error(f"Erro ao ler Excel: {e}")
            continue

        if coluna not in df.columns:
            sg.popup_error(f"A coluna '{coluna}' não existe no Excel. Colunas: {list(df.columns)}")
            continue

        urls = [str(u).strip() for u in df[coluna].dropna().tolist()]
        total = len(urls)
        if total == 0:
            sg.popup_error("Não há URLs na coluna informada.")
            continue

        os.makedirs(outdir, exist_ok=True)
        report_csv = os.path.join(outdir, "relatorio_downloads.csv")
        sess = requests_session()

        with open(report_csv, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["url", "arquivo", "metodo", "status", "bytes", "ok", "erro"])

            ok_count = 0
            for i, url in enumerate(urls, start=1):
                erro = ""
                result = None
                try:
                    # 1) tenta requests
                    result = baixar_pdf_requests(sess, url, outdir)
                except Exception as e1:
                    erro = f"requests: {e1}"
                    # 2) fallback Playwright
                    if use_pw:
                        try:
                            result = baixar_via_playwright(url, outdir, headless=headless)
                        except Exception as e2:
                            if erro:
                                erro = erro + f" | playwright: {e2}"
                            else:
                                erro = f"playwright: {e2}"

                if result:
                    writer.writerow([url, result["nome"], result["metodo"], result["status"], result["bytes"], 1, ""])
                    ok_count += 1
                    log(window, f"✅ {i}/{total} - {result['nome']} [{result['metodo']}] ({result['bytes']} bytes)")
                else:
                    writer.writerow([url, "", "", "", "", 0, erro])
                    log(window, f"❌ {i}/{total} - Falhou: {url}\n   Erro: {erro}")

                pct = int(i * 100 / total)
                window["-PB-"].update(pct)
                window.refresh()

        log(window, f"Concluído. Sucesso: {ok_count}/{total}")
        log(window, f"Relatório: {report_csv}")

        if values["-ZIP-"]:
            try:
                zip_path = os.path.join(outdir, "pdfs_baixados.zip")
                log(window, "Compactando ZIP…")
                final_zip = zip_dir(outdir, zip_path)
                log(window, f"ZIP pronto: {final_zip}")
            except Exception as e:
                log(window, f"Falha ao criar ZIP: {e}")

window.close()