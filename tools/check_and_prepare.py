"""Проверка свежей сборки перед публикацией и подготовка папки сайта.

Сравнивает dist/stats.json с тем, что сейчас опубликовано на сайте (stats.json рядом с
инструкцией). Если разделов или картинок стало заметно меньше — скорее всего, кто-то случайно
удалил кусок документа: публикацию останавливаем, на сайте остаётся прошлая версия.

    python tools/check_and_prepare.py <адрес сайта> [--force]
"""
import json
import shutil
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST, SITE = ROOT / "dist", ROOT / "site"
MAX_DROP = 0.30  # допустимая потеря разделов / картинок за одну ночь

sys.stdout.reconfigure(encoding="utf-8")
url = sys.argv[1].rstrip("/") + "/stats.json"
force = "--force" in sys.argv
new = json.loads((DIST / "stats.json").read_text(encoding="utf-8"))
print(f"Новая сборка: разделов {new['headings']}, картинок {new['images']}")

try:
    with urllib.request.urlopen(url, timeout=30) as r:
        old = json.loads(r.read().decode("utf-8"))
    print(f"На сайте сейчас: разделов {old['headings']}, картинок {old['images']} (от {old['built']})")
except Exception as e:
    old = None
    print(f"Опубликованной версии пока нет ({e}) — публикуем первую.")

if old and not force:
    for key, label in (("headings", "разделов"), ("images", "картинок")):
        if old[key] and new[key] < old[key] * (1 - MAX_DROP):
            print(f"СТОП: {label} стало {new[key]} вместо {old[key]} — похоже, часть документа удалили. "
                     "Если так и задумано, запустите публикацию вручную с галочкой «force».")
            sys.exit(1)

SITE.mkdir(exist_ok=True)
html = next(DIST.glob("*.html"))
shutil.copy(html, SITE / "index.html")
shutil.copy(next(DIST.glob("*.docx")), SITE / html.with_suffix(".docx").name)
shutil.copy(DIST / "stats.json", SITE / "stats.json")
print("Папка сайта готова:", ", ".join(p.name for p in SITE.iterdir()))
