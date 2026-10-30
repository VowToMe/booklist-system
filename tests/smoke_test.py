#!/usr/bin/env python3
"""Start the local server against the sample workbook and exercise core APIs."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample" / "booklist.sample.xlsx"


def request(url: str, payload: dict | None = None) -> dict:
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"} if body else {}
    with urllib.request.urlopen(
        urllib.request.Request(url, data=body, headers=headers), timeout=3
    ) as response:
        return json.loads(response.read().decode("utf-8"))


def main() -> int:
    if not SAMPLE.exists():
        raise SystemExit(f"missing sample workbook: {SAMPLE}")

    with tempfile.TemporaryDirectory(prefix="booklist-smoke-") as temp:
        temp_path = Path(temp)
        workbook = temp_path / "booklist.xlsx"
        shutil.copy2(SAMPLE, workbook)
        env = os.environ.copy()
        env.update(
            {
                "BOOKLIST_XLSX": str(workbook),
                "BOOKLIST_CSV": str(temp_path / "booklist.csv"),
                "BOOKLIST_DATA": str(temp_path / "data"),
                "BOOKLIST_PORT": "18770",
                "BOOKLIST_READEST": "0",
            }
        )
        process = subprocess.Popen(
            [sys.executable, str(ROOT / "server.py")],
            cwd=ROOT,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        try:
            base = "http://127.0.0.1:18770"
            for _ in range(40):
                if process.poll() is not None:
                    stdout, stderr = process.communicate()
                    raise RuntimeError(f"server exited early:\n{stdout}\n{stderr}")
                try:
                    data = request(base + "/api/books")
                    break
                except (urllib.error.URLError, TimeoutError):
                    time.sleep(0.1)
            else:
                raise RuntimeError("server did not become ready")

            books = data.get("books") or []
            assert len(books) == 5, f"expected 5 sample books, got {len(books)}"
            assert books[0]["书名"] == "The Design of Everyday Things"

            saved = request(
                base + "/api/save",
                {
                    "书名": books[0]["书名"],
                    "序号": books[0]["id"],
                    "判定": "在读",
                    "开始日": "2026-10-30",
                    "感想": "Smoke test record",
                },
            )
            assert saved["记录"]["判定"] == "在读"

            exported = request(base + "/api/export.json")
            assert exported, "JSON export is empty"
            assert any(item.get("书名") == books[0]["书名"] for item in exported.values())
            print("PASS: 5 books loaded, save round-trip and export succeeded")
            return 0
        finally:
            process.terminate()
            try:
                process.wait(timeout=3)
            except subprocess.TimeoutExpired:
                process.kill()


if __name__ == "__main__":
    raise SystemExit(main())
