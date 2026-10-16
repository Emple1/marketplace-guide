"""Сборка инструкции по модулю маркетплейсов.

Источник правды — общий Google Doc. Скрипт скачивает его в формате Word (.docx),
превращает в оформленный HTML и встраивает туда же .docx для кнопки «Скачать Word».

Запуск:
    python build.py                    # скачать Google Doc по ссылке из config.json
    python build.py путь/к/файлу.docx  # собрать из вручную скачанного .docx
    python build.py --separate-docx    # не встраивать .docx в HTML, положить рядом
    python build.py --no-compress      # не сжимать скриншоты
"""
import argparse
import base64
import datetime
import hashlib
import html
import io
import json
import re
import sys
import zipfile
from pathlib import Path

import mammoth
import requests
from bs4 import BeautifulSoup, NavigableString
from PIL import Image

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
EMBED_LIMIT_MB = 40  # больше — .docx кладём рядом, иначе HTML станет неподъёмным
MAX_SIDE = 2000      # максимальная сторона скриншота, px

sys.stdout.reconfigure(encoding="utf-8")

# Стили Google Docs / Word → HTML
STYLE_MAP = """
p[style-name='Title'] => h1.doc-title:fresh
p[style-name='title'] => h1.doc-title:fresh
p[style-name='Subtitle'] => p.doc-subtitle:fresh
p[style-name='subtitle'] => p.doc-subtitle:fresh
"""

# Абзац, начинающийся с «Слово:», превращается в цветную выноску
CALLOUTS = {
    "важно": "warn", "внимание": "warn", "обратите внимание": "warn",
    "осторожно": "danger", "нельзя": "danger",
    "совет": "tip", "подсказка": "tip", "лайфхак": "tip",
    "примечание": "info", "заметка": "info", "информация": "info",
}
CALLOUT_RE = re.compile(r"^\s*(" + "|".join(sorted(CALLOUTS, key=len, reverse=True)) + r")\s*[:!]", re.I)
ICONS = {
    "warn": '<path d="M12 9v4m0 4h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/>',
    "danger": '<circle cx="12" cy="12" r="9"/><path d="m15 9-6 6m0-6 6 6"/>',
    "tip": '<path d="M9 18h6m-5 3h4M12 3a6 6 0 0 0-3.5 10.9c.6.4 1 1.1 1 1.8V16h5v-.3c0-.7.4-1.4 1-1.8A6 6 0 0 0 12 3z"/>',
    "info": '<circle cx="12" cy="12" r="9"/><path d="M12 11v5m0-8h.01"/>',
}
LINK_ICON = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
             'stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.7 1.7"/>'
             '<path d="M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.7-1.7"/></svg>')
CAPTION_RE = re.compile(r"^\s*(рис\.|рисунок|скриншот)\s*\d*", re.I)

# Маркетплейсы для фильтра на странице: раздел относится к маркетплейсу, если тот назван
# в заголовке раздела (или выше по вложенности). Код → способы написания в заголовках.
MARKETPLACES = {
    "wb": r"wildberries|\bрвб\b|\bwb\b|\bвб\b",
    "ozon": r"ozon|озон",
    "ym": r"яндекс\s*\.?\s*маркет|\bям\b",
    "mm": r"мегамаркет|megamarket",
    "ali": r"aliexpress|алиэкспресс",
    "mvideo": r"м\.\s*видео|мвидео|эльдорадо",
    "lamoda": r"lamoda|ламода",
    "magnit": r"магнит\s*маркет",
}


def marketplaces_in(text):
    t = text.lower()
    return [code for code, rx in MARKETPLACES.items() if re.search(rx, t)]


def load_config():
    return json.loads((ROOT / "config.json").read_text(encoding="utf-8"))


def download_google_doc(doc):
    """Скачивает Google Doc как .docx. Нужен доступ «Все, у кого есть ссылка» (хотя бы чтение)."""
    m = re.search(r"/document/d/([\w-]+)", doc)
    doc_id = m.group(1) if m else doc.strip()
    url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
    print(f"Скачиваю Google Doc {doc_id} …")
    r = requests.get(url, timeout=120)
    if r.status_code != 200 or "text/html" in r.headers.get("Content-Type", ""):
        sys.exit(
            "Не удалось скачать документ: у него нет доступа по ссылке.\n"
            "Либо включите в Google Doc «Настройки доступа → Все, у кого есть ссылка (читатель)»,\n"
            "либо скачайте вручную «Файл → Скачать → Microsoft Word (.docx)» и запустите:\n"
            "    python build.py путь/к/файлу.docx"
        )
    return r.content


def doc_id(doc):
    m = re.search(r"/document/d/([\w-]+)", doc)
    return m.group(1) if m else doc.strip()


def slugify(text, used):
    s = re.sub(r"[^\w\s-]", "", text.lower(), flags=re.U)
    s = re.sub(r"[\s_-]+", "-", s).strip("-")[:60].strip("-") or "razdel"
    base, n = s, 2
    while s in used:
        s, n = f"{base}-{n}", n + 1
    used.add(s)
    return s


def norm_space(t):
    return re.sub(r"\s+", " ", t).strip()


def text_of(tag):
    return re.sub(r"\s+", " ", tag.get_text(" ")).strip()


def postprocess(fragment, own_doc_id=None):
    soup = BeautifulSoup(fragment, "html.parser")

    # Титул и подзаголовок документа показываются в шапке страницы, из текста убираем
    for t in soup.select("h1.doc-title, p.doc-subtitle"):
        t.decompose()
    # Титульный лист (всё до первого заголовка) — тоже, название уже в шапке
    first = soup.find(["h1", "h2", "h3"])
    if first is not None:
        for el in list(first.find_previous_siblings()):
            el.decompose()

    # Пустые абзацы, которые Google Docs оставляет между блоками
    for p in soup.find_all("p"):
        if not text_of(p) and not p.find("img"):
            p.decompose()

    # Ручное «Содержание» из документа не нужно — на странице есть боковое оглавление
    for h in soup.find_all("h1"):
        if text_of(h).lower() in ("содержание", "оглавление"):
            nxt = h.find_next_sibling()
            while nxt is not None and nxt.name != "h1":
                cur, nxt = nxt, nxt.find_next_sibling()
                cur.decompose()
            h.decompose()

    # «Заголовок», который целиком состоит из ссылок, — это пункты списка вопросов,
    # случайно оформленные стилем заголовка: делаем каждую ссылку обычным абзацем
    for h in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6"]):
        links = h.find_all("a", href=True)
        if links and norm_space(" ".join(text_of(a) for a in links)) == text_of(h):
            for a in links:
                p = soup.new_tag("p")
                p.append(a.extract())
                h.insert_before(p)
            for b in h.find_all("a", id=True):  # закладку заголовка переносим на первый абзац
                h.find_previous_sibling("p").insert(0, b.extract())
            h.decompose()

    # Заголовки: якоря для оглавления и ссылок «на раздел»
    used, aliases, orphans = set(), {}, []
    for h in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6"]):
        title = text_of(h)
        bookmarks = [a["id"] for a in h.find_all("a", href=False) if a.get("id")]  # закладки Google Docs
        if not title:
            orphans += bookmarks  # закладка пустого заголовка ведёт на следующий
            h.decompose()
            continue
        h["id"] = slugify(title, used)
        for b in bookmarks + orphans:
            aliases[b] = h["id"]
        orphans = []
        for a in h.find_all("a", href=False):
            a.unwrap()
        h["data-title"] = title
        anchor = soup.new_tag("a", attrs={"class": "anchor", "href": f"#{h['id']}",
                                          "title": "Скопировать ссылку на раздел"})
        anchor.append(BeautifulSoup(LINK_ICON, "html.parser"))
        h.append(anchor)

    # Внутренние ссылки документа («см. раздел …») ведут на новые якоря заголовков
    ids = {t["id"] for t in soup.find_all(id=True)}

    # Ссылки на этот же Google Doc («…/document/d/<id>/…#heading=h.xxx») уводили бы клиента
    # в исходник — превращаем их во внутренние ссылки на раздел инструкции
    if own_doc_id:
        for a in soup.find_all("a", href=lambda h: h and own_doc_id in h):
            m = re.search(r"#(?:heading=)?(h\.[\w-]+)", a["href"])
            if m:
                a["href"] = "#_heading=" + m.group(1)
                continue
            # Ссылка на документ целиком: убираем вместе с текстом-адресом
            parent = a.parent
            if text_of(a).startswith("http"):
                a.decompose()
            else:
                a.unwrap()
            if parent is not None and parent.name in ("p", "li") and len(text_of(parent)) < 80 and not parent.find("img"):
                parent.decompose()

    def norm(t):
        return re.sub(r"[\W_]+", " ", re.sub(r"\s\d+$", "", t)).strip().lower()

    def find_by_text(a):
        """Закладка потерялась — ищем раздел с тем же названием (заголовок или жирный абзац)."""
        want = norm(text_of(a))
        if not want:
            return None
        levels = ["h1", "h2", "h3", "h4", "h5", "h6"]
        below = [h for h in a.find_all_next(levels, limit=30) if h.get("id")]
        # 1) ближайший заголовок ниже с точно таким названием, 2) ниже — начинающийся с него
        #    («Wildberries» → «Wildberries - Настройка профиля»), 3) такой же в любом месте
        for h in below:
            if norm(h.get("data-title", "")) == want:
                return h["id"]
        for h in below:
            if norm(h.get("data-title", "")).startswith(want + " "):
                return h["id"]
        for h in soup.find_all(levels):
            if h.get("id") and norm(h.get("data-title", "")) == want:
                return h["id"]
        for p in soup.find_all("p"):
            if p is a.parent or p.find("a", href=True):
                continue
            st = p.find("strong")
            if st is not None and norm(text_of(p)) == want == norm(text_of(st)):
                if not p.get("id"):
                    p["id"] = slugify(text_of(p), used)
                    ids.add(p["id"])
                return p["id"]
        return None

    for a in soup.find_all("a", href=True):
        if not a["href"].startswith("#"):
            continue
        target = aliases.get(a["href"][1:], a["href"][1:])
        if target not in ids:
            target = find_by_text(a) or target
        if target in ids:
            a["href"] = "#" + target
        else:
            print(f"  · битая ссылка в документе: «{text_of(a)[:70]}» — раздел, на который она ведёт, удалён")
            a.unwrap()

    # Абзацы только из картинок → <figure>, следующая строка «Рис. …» → подпись
    for p in soup.find_all("p"):
        imgs = p.find_all("img")
        if not imgs or text_of(p):
            continue
        fig = soup.new_tag("figure")
        for img in imgs:
            img["loading"] = "lazy"
            if not img.get("alt"):
                img["alt"] = ""
            fig.append(img.extract())
        nxt = p.find_next_sibling()
        if nxt is not None and nxt.name == "p" and not nxt.find("img"):
            cap = text_of(nxt)
            whole_italic = nxt.find("em") is not None and text_of(nxt.find("em")) == cap
            if CAPTION_RE.match(cap) or (whole_italic and len(cap) < 200):
                fc = soup.new_tag("figcaption")
                for child in list(nxt.contents):
                    fc.append(child.extract())
                fig.append(fc)
                nxt.decompose()
        p.replace_with(fig)

    # Выноски «Важно: …», «Совет: …»
    for p in soup.find_all("p"):
        if p.find_parent(["table", "figure"]):
            continue
        m = CALLOUT_RE.match(text_of(p))
        if not m:
            continue
        kind = CALLOUTS[m.group(1).lower()]
        if p.find("strong") is None or not text_of(p).startswith(text_of(p.find("strong"))):
            # Выделяем слово-метку жирным, если автор этого не сделал
            first = p.contents[0] if p.contents else None
            if isinstance(first, NavigableString):
                mm = CALLOUT_RE.match(str(first))
                if mm:
                    strong = soup.new_tag("strong")
                    strong.string = str(first)[: mm.end()]
                    rest = str(first)[mm.end():]
                    first.replace_with(strong)
                    strong.insert_after(NavigableString(rest))
        box = soup.new_tag("div", attrs={"class": f"callout {kind}"})
        box.append(BeautifulSoup(
            f'<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round">{ICONS[kind]}</svg>', "html.parser"))
        body = soup.new_tag("div")
        p.replace_with(box)
        body.append(p)
        box.append(body)

    # Таблицы — с прокруткой на узком экране
    for table in soup.find_all("table"):
        if table.find_parent("table"):
            continue
        wrap = soup.new_tag("div", attrs={"class": "table-wrap"})
        table.wrap(wrap)

    # Ссылки наружу — в новой вкладке
    for a in soup.find_all("a", href=True):
        if a["href"].startswith("http"):
            a["target"] = "_blank"
            a["rel"] = "noopener"

    remove_marketplace_lists(soup)
    tag_marketplaces(soup)
    return str(soup)


def is_marketplace_line(el):
    """Абзац, в котором только название маркетплейса («Wildberries», «OZON FBS»)."""
    if el is None or el.name != "p" or el.find("img") or len(text_of(el)) >= 40:
        return False
    own = marketplaces_in(text_of(el))
    if len(own) != 1:
        return False
    rest = re.sub(MARKETPLACES[own[0]], "", text_of(el).lower())
    return not re.sub(r"[\W_]|fbs|fbo", "", rest)


def remove_marketplace_lists(soup):
    """Убирает списки-перечисления «Wildberries / OZON / ЯндексМаркет …» (3+ строки подряд):
    это навигация из Google Doc, на странице её заменяют оглавление и фильтр по маркетплейсу."""
    removed = 0
    for el in soup.find_all("p", recursive=False):
        if el.parent is None or not is_marketplace_line(el) or is_marketplace_line(el.find_previous_sibling()):
            continue
        run = [el]
        while is_marketplace_line(run[-1].find_next_sibling()):
            run.append(run[-1].find_next_sibling())
        if len(run) >= 3:
            for x in run:
                x.decompose()
            removed += 1
    if removed:
        print(f"  · убрано списков-перечислений маркетплейсов: {removed}")


def tag_marketplaces(soup):
    """Помечает блоки атрибутом data-mp="wb ozon …" — по ним работает фильтр на странице."""
    stack = []  # (уровень заголовка, маркетплейсы)
    for el in soup.find_all(recursive=False):
        if not hasattr(el, "name") or el.name is None:
            continue
        if re.fullmatch(r"h[1-6]", el.name):
            level = int(el.name[1])
            while stack and stack[-1][0] >= level:
                stack.pop()
            stack.append((level, marketplaces_in(el.get("data-title", ""))))
        mps = next((m for _, m in reversed(stack) if m), [])
        # Строка из одного названия маркетплейса («Wildberries», «OZON» в списках)
        if not mps and el.name == "p" and len(text_of(el)) < 40:
            own = marketplaces_in(text_of(el))
            rest = re.sub(MARKETPLACES[own[0]], "", text_of(el).lower()) if len(own) == 1 else "x"
            if not re.sub(r"[\W_]|fbs|fbo", "", rest):  # кроме названия ничего нет
                mps = own
        if mps:
            el["data-mp"] = " ".join(mps)


def logo_html(cfg):
    path = cfg.get("logo")
    if path and (ROOT / path).exists():
        p = ROOT / path
        mime = {"svg": "image/svg+xml", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(p.suffix[1:].lower(), "image/png")
        return f'<img src="data:{mime};base64,{base64.b64encode(p.read_bytes()).decode()}" alt="">'
    letter = html.escape((cfg.get("title") or "И")[0])
    return f'<span class="badge-logo">{letter}</span>'


def optimize_docx(docx_bytes):
    """Сжимает скриншоты внутри .docx без видимой потери качества.

    PNG → палитра 256 цветов (для скриншотов 1С разницы не видно, вес в 3–4 раза меньше),
    слишком большие картинки уменьшаются до MAX_SIDE. Размер картинок на странице Word
    задаётся в разметке документа, поэтому вёрстка Word не меняется.
    """
    src = zipfile.ZipFile(io.BytesIO(docx_bytes))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info)
            if info.filename.startswith("word/media/") and info.filename.lower().endswith(".png"):
                try:
                    im = Image.open(io.BytesIO(data))
                    im = im.convert("RGBA") if im.mode in ("P", "LA", "RGBA", "PA") else im.convert("RGB")
                    if max(im.size) > MAX_SIDE:
                        im.thumbnail((MAX_SIDE, MAX_SIDE))
                    buf = io.BytesIO()
                    im.quantize(256, method=Image.Quantize.FASTOCTREE).save(buf, format="PNG", optimize=True)
                    if buf.tell() < len(data):
                        data = buf.getvalue()
                except Exception as e:  # битая картинка — оставляем как есть
                    print(f"  · не удалось сжать {info.filename}: {e}")
            dst.writestr(info, data)
    return out.getvalue()


def media_index(docx_bytes):
    """Хэш содержимого картинки → её путь внутри .docx (word/media/imageN.png)."""
    z = zipfile.ZipFile(io.BytesIO(docx_bytes))
    return {hashlib.sha1(z.read(n)).hexdigest(): n for n in z.namelist() if n.startswith("word/media/")}


def main():
    ap = argparse.ArgumentParser(description="Сборка HTML + Word из Google Doc")
    ap.add_argument("docx", nargs="?", help="путь к .docx (иначе скачать Google Doc из config.json)")
    ap.add_argument("--separate-docx", action="store_true", help="положить .docx рядом, а не внутрь HTML")
    ap.add_argument("--no-compress", action="store_true", help="не сжимать скриншоты")
    args = ap.parse_args()
    cfg = load_config()

    if args.docx:
        docx_bytes = Path(args.docx).read_bytes()
    elif cfg.get("google_doc"):
        docx_bytes = download_google_doc(cfg["google_doc"])
    else:
        sys.exit("Укажите ссылку на Google Doc в config.json (поле google_doc) или передайте путь к .docx")

    if not args.no_compress:
        before = len(docx_bytes)
        print("Сжимаю скриншоты …")
        docx_bytes = optimize_docx(docx_bytes)
        print(f"  {before / 2**20:.1f} МБ → {len(docx_bytes) / 2**20:.1f} МБ")

    size_mb = len(docx_bytes) / 2**20
    embed = not args.separate_docx and size_mb <= EMBED_LIMIT_MB

    # Если .docx встроен в страницу, картинки не дублируем: браузер достаёт их из него сам
    media = media_index(docx_bytes) if embed else {}

    def convert_image(image):
        with image.open() as f:
            data = f.read()
        attrs = {"alt": image.alt_text or ""}
        try:  # размеры заранее — чтобы страница не «прыгала», пока грузятся картинки
            attrs["width"], attrs["height"] = (str(v) for v in Image.open(io.BytesIO(data)).size)
        except Exception:
            pass
        path = media.get(hashlib.sha1(data).hexdigest())
        if path:
            attrs["data-media"] = path
        else:
            attrs["src"] = f"data:{image.content_type};base64,{base64.b64encode(data).decode()}"
        return attrs

    print("Превращаю в HTML …")
    result = mammoth.convert_to_html(io.BytesIO(docx_bytes), style_map=STYLE_MAP,
                                     convert_image=mammoth.images.img_element(convert_image))
    for msg in result.messages:
        if "Unrecognised paragraph style" not in msg.message:
            print("  ·", msg.message)
    content = postprocess(result.value, doc_id(cfg.get("google_doc", "")))

    DIST.mkdir(exist_ok=True)
    name = cfg["file_name"]
    docx_name = f"{name}.docx"
    (DIST / docx_name).write_bytes(docx_bytes)

    if embed:
        docx_href = "#"
        docx_data = ('<script type="application/octet-stream" id="docx-data">'
                     + base64.b64encode(docx_bytes).decode() + "</script>")
    else:
        docx_href, docx_data = docx_name, ""

    today = datetime.date.today().strftime("%d.%m.%Y")
    page = (ROOT / "template.html").read_text(encoding="utf-8")
    for key, value in {
        "TITLE": html.escape(cfg["title"]),
        "KICKER": html.escape(cfg.get("kicker", "")),
        "VERSION": html.escape(str(cfg.get("version", "1.0"))),
        "BUILD_DATE": today,
        "ACCENT": cfg.get("accent", "#2563eb"),
        "FOOTER": html.escape(cfg.get("footer", "")),
        "LOGO": logo_html(cfg),
        "DOCX_NAME": html.escape(docx_name),
        "DOCX_HREF": html.escape(docx_href),
    }.items():
        page = page.replace("{{" + key + "}}", value)
    # Содержимое подставляем последним, чтобы случайные {{…}} в тексте не трогать
    page = page.replace("{{DOCX_DATA}}", docx_data).replace("{{CONTENT}}", content)

    out = DIST / f"{name}.html"
    out.write_text(page, encoding="utf-8")
    # Сводка для проверки перед публикацией: не пропала ли половина документа
    stats = {"headings": content.count("<h1") + content.count("<h2") + content.count("<h3"),
             "images": content.count("<img"), "built": today}
    (DIST / "stats.json").write_text(json.dumps(stats, ensure_ascii=False), encoding="utf-8")
    print(f"Готово: {out}  ({out.stat().st_size / 2**20:.1f} МБ)")
    print(f"        {DIST / docx_name}  ({size_mb:.1f} МБ)"
          + ("" if embed else "  ← отправляйте клиенту вместе с HTML"))


if __name__ == "__main__":
    main()
