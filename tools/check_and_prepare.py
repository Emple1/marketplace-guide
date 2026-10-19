"""Проверка свежей сборки перед публикацией и подготовка папки сайта.

Сравнивает dist/stats.json с тем, что сейчас опубликовано на сайте (stats.json рядом с
инструкцией). Если разделов или картинок стало заметно меньше — скорее всего, кто-то случайно
удалил кусок документа: публикацию останавливаем, на сайте остаётся прошлая версия.

Заодно ведёт changes.json для блока «Что нового»: сравнивает отпечатки разделов (sections.json)
новой сборки с опубликованными и дописывает добавленные / обновлённые разделы за последние 14 дней.

    python tools/check_and_prepare.py <адрес сайта> [--force]
"""
import datetime
import json
import shutil
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST, SITE = ROOT / "dist", ROOT / "site"
MAX_DROP = 0.30  # допустимая потеря разделов / картинок за одну ночь
KEEP_DAYS = 14   # сколько дней раздел показывается в «Что нового»
MAX_CHANGED = 0.40  # изменилось больше — это перестройка документа, а не правки: в «Что нового» не пишем

sys.stdout.reconfigure(encoding="utf-8")
base = sys.argv[1].rstrip("/")
force = "--force" in sys.argv
new = json.loads((DIST / "stats.json").read_text(encoding="utf-8"))
print(f"Новая сборка: разделов {new['headings']}, картинок {new['images']}")


def fetch_json(name):
    with urllib.request.urlopen(f"{base}/{name}", timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch_optional(name, default):
    try:
        return fetch_json(name)
    except Exception:
        return default


def whats_new():
    """Добавленные / обновлённые разделы за KEEP_DAYS дней → список для changes.json."""
    today = datetime.date.today()
    changes = [c for c in fetch_optional("changes.json", [])
               if (today - datetime.date.fromisoformat(c["date"])).days < KEEP_DAYS]
    cur = json.loads((DIST / "web" / "sections.json").read_text(encoding="utf-8"))
    prev = fetch_optional("sections.json", None)
    if prev is None:  # первая сборка с отпечатками — сравнивать не с чем
        print("Что нового: отпечатков разделов на сайте ещё нет — начинаю отсчёт с этой сборки.")
        return changes
    prev = {s["id"]: s["hash"] for s in prev}
    fresh = [{"date": today.isoformat(), "id": s["id"], "title": s["title"],
              "kind": "added" if s["id"] not in prev else "updated"}
             for s in cur if prev.get(s["id"]) != s["hash"]]
    if len(fresh) > len(cur) * MAX_CHANGED:
        print(f"Что нового: изменилось {len(fresh)} из {len(cur)} разделов — похоже на перестройку документа, "
              "в список не добавляю.")
        return changes
    ids = {c["id"] for c in fresh}
    changes = fresh + [c for c in changes if c["id"] not in ids]
    print(f"Что нового: сегодня {len(fresh)} разделов, всего в списке {len(changes)}.")
    return changes


try:
    old = fetch_json("stats.json")
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

shutil.rmtree(SITE, ignore_errors=True)
shutil.copytree(DIST / "web", SITE)
(SITE / "changes.json").write_text(json.dumps(whats_new(), ensure_ascii=False), encoding="utf-8")
print("Папка сайта готова:", len(list(SITE.rglob("*"))), "файлов")
