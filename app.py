# trecho utilitário Playwright (Python)
from playwright.sync_api import sync_playwright
import os, time
from urllib.parse import urlparse

def baixar_via_playwright(url, pasta_destino, timeout_ms=90000):
    os.makedirs(pasta_destino, exist_ok=True)

    # Dica: com PLAYWRIGHT_BROWSERS_PATH=0 no build, os browsers ficam embalados no app.
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)  # headless False para depurar
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # 1) tenta capturar um download (muitos portais disparam o PDF como "download" direto)
        got_download = None
        def _on_download(download):
            nonlocal got_download
            got_download = download

        page.on("download", _on_download)
        page.goto(url, timeout=timeout_ms, wait_until="domcontentloaded")

        # Heurística: se a página tiver link explícito para PDF, clica
        # (ajuste o seletor conforme sua realidade, ou mantenha só a navegação)
        try:
            page.click("a[href$='.pdf'], a[href*='.pdf?']", timeout=3000)
            # O clique pode disparar o evento download; aguardamos um pouco
            for _ in range(30):
                if got_download:
                    break
                time.sleep(0.2)
        except:
            pass

        if got_download:
            # salva usando o nome sugerido
            sug = got_download.suggested_filename or "arquivo.pdf"
            destino = os.path.join(pasta_destino, sug)
            got_download.save_as(destino)
            context.close(); browser.close()
            return sug, destino

        # 2) Se não houve evento de download, salvamos a "page" atual,
        # que às vezes é o PDF renderizado (Chrome PDF viewer) – baixamos o conteúdo bruto.
        # Observação: nem todos servidores permitem; trate exceções no seu fluxo.
        # Abaixo tentamos pegar o conteúdo pela resposta principal:
        main_resp = page.main_frame.response
        if main_resp and ("pdf" in (main_resp.headers.get("content-type","").lower())):
            nome = os.path.basename(urlparse(page.url).path) or "arquivo.pdf"
            if not nome.lower().endswith(".pdf"):
                nome += ".pdf"
            destino = os.path.join(pasta_destino, nome)
            # Baixa o corpo da resposta (bytes)
            body = main_resp.body()
            with open(destino, "wb") as f:
                f.write(body)
            context.close(); browser.close()
            return nome, destino

        context.close(); browser.close()
        raise RuntimeError("Não foi possível capturar o PDF automaticamente nesta URL.")