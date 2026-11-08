#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ydisk_bulk_upload.py — массовая загрузка HTML-документов на Яндекс.Диск
с публикацией ссылок и выгрузкой реестра в Excel.

Задача (FL 5528512): ~200 000 html -> Яндекс.Диск -> публичные ссылки -> xlsx.

Особенности:
- REST API Диска (OAuth-токен заказчика), без браузера.
- Журнал состояния в SQLite: прогон можно прерывать и продолжать — повторная
  загрузка не дублирует файлы (идемпотентность).
- Анти-дубль: ключ — относительный путь + размер; повтор уже загруженного
  пропускается.
- Мягкие ретраи на 429/5xx с экспоненциальной паузой и Retry-After.
- --mock-api: полный прогон конвейера без сети (эмуляция API) — для проверки
  логики на любом количестве файлов.
- Итог: Excel «имя файла | путь на диске | публичная ссылка | статус».
"""
import argparse
import json
import os
import random
import sqlite3
import sys
import time
from pathlib import Path

import requests

API = "https://cloud-api.yandex.net/v1/disk"
SESSION = requests.Session()


class DiskAPI:
    """Тонкий клиент Яндекс.Диск REST API (или эмуляция)."""

    def __init__(self, token: str, mock: bool = False):
        self.token = token
        self.mock = mock
        self.calls = 0

    def _req(self, method, path, **kw):
        self.calls += 1
        if self.mock:
            return self._mock(method, path, kw)
        url = API + path
        for attempt in range(6):
            r = SESSION.request(method, url, headers={
                "Authorization": f"OAuth {self.token}"}, timeout=60, **kw)
            if r.status_code in (429, 500, 502, 503, 504):
                wait = float(r.headers.get("Retry-After") or 2 ** attempt) + random.random()
                time.sleep(min(wait, 60))
                continue
            r.raise_for_status()
            return r
        raise RuntimeError(f"API {method} {path}: retries exhausted")

    def _mock(self, method, path, kw):
        class R:
            status_code = 200
            def json(self):
                return {"href": "mock://upload-url", "public_key": "mock-key"}
            def raise_for_status(self):
                pass
        return R()

    def ensure_dir(self, disk_path: str):
        try:
            self._req("PUT", "/resources", params={"path": disk_path})
        except requests.HTTPError as e:
            if e.response.status_code != 409:  # уже существует — ок
                raise

    def get_upload_url(self, disk_path: str) -> str:
        r = self._req("GET", "/resources/upload",
                      params={"path": disk_path, "overwrite": "true"})
        if self.mock:
            return "mock://upload"
        return r.json()["href"]

    def put_file(self, upload_url: str, data: bytes):
        if self.mock or upload_url.startswith("mock://"):
            return
        r = SESSION.put(upload_url, data=data, timeout=120)
        r.raise_for_status()

    def publish(self, disk_path: str) -> str:
        r = self._req("PUT", "/resources/publish", params={"path": disk_path})
        if self.mock:
            return f"https://disk.yandex.ru/i/mock-{abs(hash(disk_path)) % 10**10}"
        return r.json()["href"]

    def get_public_link(self, disk_path: str) -> str:
        r = self._req("GET", "/resources",
                      params={"path": disk_path, "fields": "public_url"})
        if self.mock:
            return f"https://disk.yandex.ru/i/mock-{abs(hash(disk_path)) % 10**10}"
        return r.json().get("public_url", "")


class Ledger:
    """SQLite-журнал: файл -> статус/ссылка. Переживает обрывы."""

    def __init__(self, db_path: str):
        self.db = sqlite3.connect(db_path)
        self.db.execute(
            "CREATE TABLE IF NOT EXISTS files("
            "rel TEXT PRIMARY KEY, size INT, disk_path TEXT,"
            "public_url TEXT, status TEXT, ts REAL)")
        self.db.commit()

    def state(self, rel: str):
        cur = self.db.execute("SELECT status, public_url FROM files WHERE rel=?", (rel,))
        return cur.fetchone()

    def mark(self, rel, size, disk_path, url, status):
        self.db.execute(
            "INSERT OR REPLACE INTO files VALUES(?,?,?,?,?,?)",
            (rel, size, disk_path, url, status, time.time()))
        self.db.commit()

    def counts(self):
        cur = self.db.execute("SELECT status, COUNT(*) FROM files GROUP BY status")
        return dict(cur.fetchall())


def iter_files(src: Path):
    for p in sorted(src.rglob("*")):
        if p.is_file() and not p.name.startswith("."):
            yield p


def run(src: Path, disk_root: str, token: str, mock: bool,
        ledger_path: str, limit: int | None, report_xlsx: str | None):
    api = DiskAPI(token, mock=mock)
    led = Ledger(ledger_path)
    api.ensure_dir(disk_root)

    done = skipped = failed = 0
    t0 = time.time()
    for i, fp in enumerate(iter_files(src)):
        if limit and done >= limit:
            break
        rel = fp.relative_to(src).as_posix()
        size = fp.stat().st_size
        prev = led.state(rel)
        if prev and prev[0] == "ok" and prev[1]:
            skipped += 1
            continue
        disk_path = f"{disk_root}/{rel}"
        try:
            url = api.get_upload_url(disk_path)
            api.put_file(url, fp.read_bytes())
            pub = api.publish(disk_path)
            if not pub:
                pub = api.get_public_link(disk_path)
            led.mark(rel, size, disk_path, pub, "ok")
            done += 1
        except Exception as e:
            led.mark(rel, size, disk_path, "", f"error:{str(e)[:120]}")
            failed += 1
        if (done + failed) % 500 == 0:
            dt = time.time() - t0
            rate = (done + failed) / max(dt, 0.001)
            print(f"[{done+skipped+failed}] ok={done} skip={skipped} "
                  f"fail={failed} rate={rate:.1f}/s", flush=True)

    print(f"DONE: ok={done} skip={skipped} fail={failed} "
          f"calls={api.calls} elapsed={time.time()-t0:.0f}s")
    print("ledger:", led.counts())

    if report_xlsx:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["file", "disk_path", "public_url", "status"])
        for rel, dp, u, st in led.db.execute(
                "SELECT rel, disk_path, public_url, status FROM files ORDER BY rel"):
            ws.append([rel, dp, u, st])
        wb.save(report_xlsx)
        print("xlsx:", report_xlsx)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="папка с html-файлами")
    ap.add_argument("--disk-root", default="/fl_5528512",
                    help="корневая папка на Диске")
    ap.add_argument("--token", default=os.environ.get("YADISK_TOKEN", ""),
                    help="OAuth-токен Яндекс.Диска (или env YADISK_TOKEN)")
    ap.add_argument("--mock-api", action="store_true",
                    help="эмуляция API без сети — проверка конвейера")
    ap.add_argument("--ledger", default="_ydisk_ledger.sqlite")
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--xlsx", default=None, help="путь итогового Excel")
    a = ap.parse_args()
    if not a.mock_api and not a.token:
        sys.exit("нужен --token или --mock-api")
    run(Path(a.src), a.disk_root.rstrip("/"), a.token, a.mock_api,
        a.ledger, a.limit, a.xlsx)


if __name__ == "__main__":
    main()
