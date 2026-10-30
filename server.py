r"""书单 · 本地服务(默认 127.0.0.1:8770, v2)

底本(只读): booklist.xlsx
我的标记:   data\我的标记.json  —— 判定 / 感想 / 开始日 / 读完日 / 清单
            主键 = 底本「序号」, 键值里同时存 书名; 认领顺序: 书名 → 归一化书名 → 序号
            三条认领通路都要过「指纹」: 记录里的 作者/加入日期 与底本这一行对不上就不认,
            退到 orphans(界面上的「已不在底本」), 宁可显式变成孤儿, 也不静默挂到别的书上。
            名字/序号算不出来的记录进 orphans, 界面上可以「挂到某本书」或「删掉这条」。
清单:       一本书可以同时在多个清单里(多对多, 对应别的项目里的 listIds / 书架中间表)。
            底本备注里写的「主题：X」会自动并进 清单, 但只有网页上加的才撤得掉。

接口
  GET  /                  单页应用
  GET  /favicon.ico       图标
  GET  /api/books         书 + 我的标记(合并后) + 孤儿记录
  GET  /api/store         原始标记 JSON(排查用)
  GET  /api/export.md     导出 Markdown(带日期)
  GET  /api/export.json   导出 JSON
  POST /api/save          {书名, 序号?, 判定?, 感想?, 开始日?, 读完日?, 清单?} → {记录, 书名}
                          页面过期(底本的序号对不上/书名找不到)时回 409, 让页面刷新, 不瞎写
  POST /api/bulk          {items:[{书名, 序号?}...], 判定?/开始日?/读完日?}
                          → {ok, 改动:[{书名, 前, 后}]}
                          批量打状态: 整单一起校验、整单一起写(要么全成, 要么一个都不动)。
                          「前」= 这次动过的那几个字段改之前的样子, 页面靠它做整批撤销。
  POST /api/relink        {书名, 到} → 把孤儿记录挂到另一本书上

环境变量(自检用, 不影响正常启动)
  BOOKLIST_PORT / BOOKLIST_DATA / BOOKLIST_XLSX / BOOKLIST_CSV
"""
import csv
import datetime
import io
import json
import os
import re
import shutil
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX = os.environ.get("BOOKLIST_XLSX") or os.path.join(ROOT, "booklist.xlsx")
CSV = os.environ.get("BOOKLIST_CSV") or os.path.join(ROOT, "booklist.csv")
DATA_DIR = os.environ.get("BOOKLIST_DATA") or os.path.join(ROOT, "data")
STORE = os.path.join(DATA_DIR, "我的标记.json")
COVER_DIR = os.path.join(DATA_DIR, "covers")
COVER_INDEX = os.path.join(DATA_DIR, "covers.json")
PORT = int(os.environ.get("BOOKLIST_PORT") or 8770)

# Readest 的本地数据(只读)。0.12.x 把每次翻页写进 KOReader 同构的 statistics.db:
#   page_stat_data(id_book, page, start_time, duration, total_pages)
# 书单只读它来算「读过几天 / 多久 / 什么时候开始和读完」, 不碰 Readest 的任何文件。
READEST_DIR = os.environ.get("BOOKLIST_READEST_DIR") or os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"), "com.bilingify.readest", "Readest")
READEST_DB = os.path.join(READEST_DIR, "statistics.db")
READEST_LIB = os.path.join(READEST_DIR, "Books", "library.json")
AUTO_FILE = os.path.join(DATA_DIR, "readest-auto.json")
READEST_DAY_MIN = 60          # 一天累计满 60 秒才算「读过一天」

# 主题别名: 底本备注里写「主题：决策与判断」, 网页上显示成「决策能力」
TOPIC_ALIAS = {"决策与判断": "决策能力"}

# 判定词表(顺序即界面顺序)。'' = 未标
VERDICTS = ("", "想读", "在读", "读完", "踢掉")
# 「读完」之后才问的那个问题: 这本书你喜不喜欢。'' = 还没说
EVALS = ("", "好看", "一般", "不好看")
# 老版本数据里的写法 → 新写法
VERDICT_LEGACY = {"留": "想读", "踢": "踢掉", "读中": "在读"}
# 清单: 一次最多挂这么多个, 每个名字最多这么长(界面上是一排 chips, 不给人写小作文)
LIST_MAX = 20
LIST_NAME_MAX = 24

DATE_RE = re.compile(r"^\d{4}(-\d{2}(-\d{2})?)?$")
_lock = threading.Lock()
MAX_BODY = 5_000_000
_books_cache = {"key": None, "books": [], "src": ""}

# 合并视图里允许被记录覆盖的字段(作者/加入日期只用来做「指纹」, 不许被记录冲掉)
REC_SHOW = ("判定", "评价", "感想", "开始日", "读完日", "阅读记录", "更新时间")
# 没有记录时给底本行铺的空白。阅读记录是列表, 单独铺([]), 不能跟字符串挤在一个 dict 里
BLANK_SHOW = {"判定": "", "评价": "", "感想": "", "开始日": "", "读完日": "", "更新时间": ""}


class StoreError(Exception):
    """标记文件读不出来。绝不当成「空的」继续跑 —— 那会在下一次保存时把原记录整个盖掉。"""


def date_ok(s):
    """2026 / 2026-09 / 2026-09-28 三种精度, 且必须是真实存在的日历日。"""
    s = str(s or "").strip()
    if not DATE_RE.match(s):
        return False
    p = [int(x) for x in s.split("-")]
    if not (1900 <= p[0] <= 2100):
        return False
    if len(p) < 2:
        return True
    if not (1 <= p[1] <= 12):
        return False
    if len(p) < 3:
        return True
    dim = (31, 29 if (p[0] % 4 == 0 and (p[0] % 100 or p[0] % 400 == 0)) else 28, 31, 30, 31, 30,
           31, 31, 30, 31, 30, 31)[p[1] - 1]
    return 1 <= p[2] <= dim


def norm_title(s):
    """书名归一化: 去掉书名号/空白/大小写差异, 用于兜底匹配被改过名的记录。"""
    return re.sub(r"[\s《》〈〉「」『』\"'“”‘’]", "", str(s or "")).lower()


def load_cover_index():
    """封面索引只是视图层缓存, 坏了就退回占位封, 绝不能拖垮书单本体。"""
    try:
        with io.open(COVER_INDEX, "r", encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict) and isinstance(d.get("books"), dict):
            return d
    except (OSError, ValueError, TypeError):
        pass
    return {"version": 1, "generated_at": "", "books": {}}


def save_cover_index(d):
    os.makedirs(COVER_DIR, exist_ok=True)
    if os.path.exists(COVER_INDEX):
        try:
            shutil.copyfile(COVER_INDEX, COVER_INDEX + ".bak")
        except OSError:
            pass
    tmp = COVER_INDEX + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="\n") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tmp, COVER_INDEX)


def cover_brief(entry, book):
    """给 /api/books 的最小封面对象; 候选图不上首屏, 免得 153 本一次拉一堆路径。"""
    entry = entry if isinstance(entry, dict) else {}
    state = str(entry.get("state") or "placeholder")
    if state not in ("ready", "review", "placeholder"):
        state = "placeholder"
    selected = entry.get("selected") if isinstance(entry.get("selected"), dict) else {}
    ready = state == "ready" and bool(selected.get("full"))
    return {
        "thumb": str(selected.get("thumb") or "") if ready else "",
        "full": str(selected.get("full") or "") if ready else "",
        "state": "ready" if ready else state,
        "source": str(selected.get("source") or entry.get("source") or ""),
        "confidence": round(float(selected.get("confidence") or entry.get("confidence") or 0), 3),
        "isbn": str(selected.get("isbn") or entry.get("isbn") or ""),
        "candidate_count": len(entry.get("candidates") or []),
        "placeholder": {
            "title": str(book.get("书名") or ""),
            "author": str(book.get("作者") or ""),
            "category": str(book.get("类别") or ""),
            "id": str(book.get("id") or ""),
        },
    }


def parse_topic(note):
    m = re.search(r"主题[：:]\s*([^（(。;；\n]+)", note or "")
    return (m.group(1).strip() if m else "") or "未分类"


def clean_lists(v):
    """任一来源的「清单」→ 一串干净的名字。不合法直接抛(写盘前整单拒绝)。"""
    if v is None:
        return []
    if isinstance(v, str):                     # 单名字也收, 免得前端少包一层数组就报错
        v = [v]
    if not isinstance(v, (list, tuple)):
        raise ValueError("清单 must be a list of names")
    out = []
    for x in v:
        s = str(x or "").strip()
        if not s:
            continue
        s = s[:LIST_NAME_MAX]
        if s not in out:
            out.append(s)
        if len(out) >= LIST_MAX:
            break
    return out


def merge_lists(topic, own):
    """底本备注里的「主题：X」+ 网页上加的清单, 去重保序。

    主题是底本的账, 清单是网页的账; 合起来才是「这本书现在属于哪几栏」。
    """
    out = []
    if topic and topic != "未分类":
        out.append(topic)
    for x in own or []:
        if x not in out:
            out.append(x)
    return out


def renames(st):
    """store 顶层那份「重命名」映射(原始名 → 界面上显示的名字), 只留真有意义的项。

    底本备注里的「主题：X」改不了(xlsx 只读), 所以改名只能记成一张映射表,
    在视图层翻一遍 —— 底本一个字节不碰, 数据也不会因为改名丢。
    """
    r = st.get("重命名") if isinstance(st, dict) else None
    out = {}
    if isinstance(r, dict):
        for k, v in r.items():
            k, v = str(k or "").strip(), str(v or "").strip()
            if k and v and k != v:
                out[k] = v
    return out


def disp(name, rn):
    """原始清单名 → 界面上显示的名字。没改过名的原样返回。"""
    return rn.get(name, name) if name else name


def out_rec(books, rec, rn=None):
    """回给前端的记录: 清单按视图口径跟底本「主题」合并, 再用改名表翻一遍。"""
    rn = rn or {}
    out = dict(rec)
    out["阅读记录"] = copy_log(rec.get("阅读记录"))
    topic = ""
    if rec.get("序号"):
        for b in books:
            if b["id"] and b["id"] == rec["序号"]:
                topic = disp(b["主题"], rn)
                break
    out["清单"] = merge_lists(topic, [disp(n, rn) for n in (rec.get("清单") or [])])
    return out


LOG_MAX = 2000          # 一本书最多留多少条「哪天读了」
LOG_PAGE_MAX = 40       # 「读到哪」那一格
LOG_NOTE_MAX = 300      # 每条最多一句话


def _log_entry(raw):
    """一条阅读记录 → 干净的一条; 日期不是完整真实日历日就返回 None。

    这里要求「完整日期」(2026-09-28): 开始日/读完日可以是「2026-09」这种粗粒度,
    但「哪一天读了」本身就是这件事的全部信息, 只到月份就没有意义了。
    """
    if not isinstance(raw, dict):
        return None
    d = str(raw.get("日期") or "").strip()
    if len(d) != 10 or not date_ok(d):
        return None
    return {"日期": d,
            "页码": str(raw.get("页码") or "").strip()[:LOG_PAGE_MAX],
            "备注": str(raw.get("备注") or "").strip()[:LOG_NOTE_MAX]}


def clean_readlog(v, strict=False):
    """阅读记录 → 按日期从新到旧排好的一串; 同一天记多条, 按你记的顺序挨着排。

    strict=True(写入路径): 有一条不合规就整单拒绝, 不留半截。
    strict=False(读旧文件): 坏的那几条丢掉, 不让整个文件读不出来。
    """
    if v is None:
        return []
    if not isinstance(v, (list, tuple)):
        if strict:
            raise ValueError("阅读记录 must be a list")
        return []
    out, seen = [], set()
    for raw in v:
        e = _log_entry(raw)
        if e is None:
            if strict:
                raise ValueError("阅读记录里每一条都要有真实日期(YYYY-MM-DD): %s" % (raw,))
            continue
        k = (e["日期"], e["页码"], e["备注"])
        if k in seen:                      # 同一秒连点两下不该留两条一样的
            continue
        seen.add(k)
        out.append(e)
    out.sort(key=lambda e: e["日期"], reverse=True)   # 排序稳定: 同一天的保持记录顺序
    return out[:LOG_MAX]


def copy_log(v):
    """把阅读记录拷一份再交出去, 免得调用方就地改到 store 里的那个 list。"""
    return [dict(x) for x in (v or [])]


def load_books():
    """底本 → 书列表。xlsx 优先, 失败回落 CSV。按 mtime+size 缓存, 底本改了自动失效。"""
    def stamp(fp):
        try:
            st = os.stat(fp)
            return (int(st.st_mtime_ns), st.st_size)
        except OSError:
            return None

    key = (stamp(XLSX), stamp(CSV))
    if key == _books_cache["key"] and _books_cache["books"]:
        return _books_cache["books"], _books_cache["src"]

    rows, src = [], "xlsx"
    try:
        import openpyxl
        wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        head = [str(h).strip() if h is not None else "" for h in next(it)]
        for r in it:
            if not r or len(r) < 2 or not r[1]:
                continue
            rows.append({head[i]: r[i] for i in range(min(len(head), len(r)))})
        wb.close()
    except Exception as e:
        src = "csv-fallback(%s)" % e
        if os.path.exists(CSV):
            with io.open(CSV, "r", encoding="utf-8-sig", newline="") as f:
                rows = [dict(x) for x in csv.DictReader(f)]

    def g(d, k):
        v = d.get(k)
        return "" if v is None else str(v).strip()

    books = []
    for d in rows:
        name = g(d, "书名")
        if not name:
            continue
        note = g(d, "备注")
        topic = parse_topic(note)
        books.append({
            "id": g(d, "序号"), "书名": name, "作者": g(d, "作者"),
            "类别": g(d, "类别"), "优先级": g(d, "优先级"),
            "来源": g(d, "来源"), "加入日期": g(d, "加入日期"),
            "为什么加": g(d, "为什么加"), "备注": note,
            "主题": TOPIC_ALIAS.get(topic, topic),
        })
    _books_cache.update({"key": key, "books": books, "src": src})
    return books, src


def rebuild_csv():
    """xlsx → CSV 镜像。底本一改就要跟着重写, 否则别的工具读到的是旧表。"""
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    with io.open(CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        for r in rows:
            w.writerow(["" if v is None else v for v in r])


def base_delete(names, ids=()):
    """从底本 xlsx 里真删掉这些书: 备份 → 删行 → 重排序号 → 修筛选 → 重建 CSV。

    返回 (删掉的书名列表, 备份文件名)。底本被 Excel/WPS 占着时 openpyxl 存盘会抛
    PermissionError, 调用方必须整单放弃 —— 不能「记录删了、书还在」, 那正是原来的毛病。
    """
    import openpyxl
    want = {norm_title(n) for n in names if str(n or "").strip()}
    want_ids = {str(i).strip() for i in ids if str(i or "").strip()}
    wb = openpyxl.load_workbook(XLSX)
    ws = wb[wb.sheetnames[0]]
    head = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    if "书名" not in head or "序号" not in head:
        wb.close()
        raise ValueError("底本第一行里找不到「书名」或「序号」")
    i_name, i_id = head.index("书名") + 1, head.index("序号") + 1
    gone, rows = [], []
    for r in range(2, ws.max_row + 1):
        raw = ws.cell(row=r, column=i_name).value
        nm = "" if raw is None else str(raw).strip()
        if not nm:
            continue
        no = ws.cell(row=r, column=i_id).value
        no = "" if no is None else str(no).strip()
        if norm_title(nm) in want or (no and no in want_ids):
            gone.append(nm)
            rows.append(r)
    if not rows:
        wb.close()
        return [], ""
    stamp = time.strftime("%Y%m%d-%H%M%S")
    stem, ext = os.path.splitext(XLSX)
    bak = "%s.bak-%s%s" % (stem, stamp, ext)
    shutil.copy2(XLSX, bak)
    for r in reversed(rows):
        ws.delete_rows(r)
    for i, r in enumerate(range(2, ws.max_row + 1), start=1):
        ws.cell(row=r, column=i_id, value=i)
    last = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[min(len(head), 26) - 1]
    ws.auto_filter.ref = "A1:%s%d" % (last, max(ws.max_row, 1))
    try:
        wb.save(XLSX)
    except PermissionError:
        wb.close()
        raise
    wb.close()
    _books_cache["key"] = None          # 让合并视图下次重新读表
    rebuild_csv()
    return gone, os.path.basename(bak)


def resync_ids(st, books):
    """删行之后序号整体前移: 把标记里的「序号」跟着改, 免得指纹落到别的书上。"""
    by_name = {b["书名"]: b["id"] for b in books}
    by_norm = {norm_title(b["书名"]): b["id"] for b in books}
    for k, v in st["items"].items():
        nb = by_name.get(k) or by_norm.get(norm_title(k))
        if nb and str(v.get("序号") or "") != nb:
            v["序号"] = nb


def blank_rec():
    r = {"序号": "", "作者": "", "加入日期": "", "清单": [], "阅读记录": []}
    r.update(BLANK_SHOW)
    return r


def clean_rec(raw):
    """把任意来源的记录规整成 v2 形状(含老词表映射)。"""
    r = blank_rec()
    if isinstance(raw, dict):
        v = str(raw.get("判定") or "").strip()
        r["判定"] = VERDICT_LEGACY.get(v, v if v in VERDICTS else "")
        e = str(raw.get("评价") or "").strip()
        r["评价"] = e if e in EVALS else ""
        r["感想"] = str(raw.get("感想") or "")
        r["序号"] = str(raw.get("序号") or "")
        r["作者"] = str(raw.get("作者") or "").strip()
        r["加入日期"] = str(raw.get("加入日期") or "").strip()
        try:
            r["清单"] = clean_lists(raw.get("清单"))
        except ValueError:
            r["清单"] = []                     # 老数据里写坏了就当没有, 不让整个文件读不出来
        r["阅读记录"] = clean_readlog(raw.get("阅读记录"))   # 同上: 坏的那几条丢掉, 不整个文件报废
        for k in ("开始日", "读完日"):
            d = str(raw.get(k) or "").strip()
            r[k] = d if date_ok(d) else ""
        r["更新时间"] = str(raw.get("更新时间") or "")
    return r


def is_empty(r):
    return not (r["判定"] or r.get("评价") or r["感想"].strip() or r["开始日"] or r["读完日"]
                or r.get("清单") or r.get("阅读记录"))


def fp_ok(book, rec):
    """这道记录还属不属于这本书?

    比对「记录里存的指纹」和「底本这一行的指纹」;两边都有的项必须全对上。
    没有可比项(老记录没存指纹 / 底本这两列本来就空)才退回旧行为(认)。
    宁可判成孤儿(界面看得见、能手动挂回去),也不把感想静默挂到别的书上。
    """
    for rec_v, book_v in ((rec.get("作者"), book.get("作者")),
                          (rec.get("加入日期"), book.get("加入日期"))):
        if rec_v and book_v and str(rec_v).strip() != str(book_v).strip():
            return False
    return True


def load_store(quarantine=True):
    """quarantine=False 给后台自动同步用: 只读一眼, 坏文件留给真正要报错的那条路去留证。

    否则同步先把文件改成 .bad 再 swallow 掉异常, /api/books 随后看到的是一份「不存在」,
    会误报 200 + 空标记 —— 用户看到所有判定凭空消失, 却没有任何提示。
    """
    def bad(why):
        if quarantine:
            return StoreError(_quarantine(why))
        return StoreError("标记文件读不出来: %s" % why)

    if not os.path.exists(STORE):
        return {"version": 2, "saved_at": "", "items": {}}
    try:
        with io.open(STORE, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception as e:
        raise bad(e)
    if not isinstance(d, dict):
        raise bad("顶层不是 JSON 对象, 而是 %s" % type(d).__name__)
    if "items" not in d:                      # 更早的写法: 顶层直接是 items
        if not all(isinstance(v, dict) for v in d.values()):
            raise bad("既没有 items 字段, 顶层也不像记录表")
        d = {"version": 2, "saved_at": "", "items": d}
    if not isinstance(d.get("items"), dict):
        raise bad("items 不是对象")
    d["version"] = 2
    d["items"] = {k: clean_rec(v) for k, v in (d.get("items") or {}).items()}
    return d


def _quarantine(why):
    """坏掉的标记文件改名留证(不删), 然后带着原因抛错。"""
    bad = STORE + ".bad-" + time.strftime("%Y%m%d-%H%M%S")
    try:
        os.replace(STORE, bad)
        where = os.path.basename(bad)
    except OSError:
        where = "(改名失败, 文件原地没动)"
    return "标记文件读不出来: %s —— 已留证为 %s, 没有覆盖它" % (why, where)


def save_store(d):
    """原子写: 先写同目录 tmp, 再 os.replace; 覆盖之前留一份 .bak。"""
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(STORE):
        try:
            with open(STORE, "rb") as src:
                data = src.read()
            with open(STORE + ".bak", "wb") as dst:
                dst.write(data)
        except OSError:
            pass                                  # 备份失败不影响正事
    tmp = STORE + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="\n") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STORE)


# ── Readest 阅读统计(只读接入) ──────────────────────────────
# Readest 每翻一页记一条事件: 第几页、从哪一秒开始、停留多少秒。
# 这里的口径: 一天累计 ≥60 秒才算读过一天; 开始/最近取这些「读书日」;
# 读完 = 第一次翻到最后一页, 或 Readest 自己标了 finished 的那天(取较早的)。
_rd_cache = {"key": None, "raw": None}


def _stamp(fp):
    try:
        st = os.stat(fp)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _day(ts):
    return time.strftime("%Y-%m-%d", time.localtime(int(ts)))


def _load_readest_lib():
    """hash → library.json 里的那条书(标题/作者/读完状态/进度)。读不出来就当没有。"""
    try:
        with io.open(READEST_LIB, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        return {}
    out = {}
    if isinstance(raw, list):
        for b in raw:
            if isinstance(b, dict) and b.get("hash"):
                out[str(b["hash"])] = b
    return out


def readest_raw():
    """Readest 统计库 → 每本书的聚合。按 db / wal / library.json 的 mtime 缓存。"""
    key = (_stamp(READEST_DB), _stamp(READEST_DB + "-wal"), _stamp(READEST_LIB))
    if key == _rd_cache["key"] and _rd_cache["raw"] is not None:
        return _rd_cache["raw"]
    snap = {"available": False, "computed_at": int(time.time()), "books": []}
    if not key[0]:
        _rd_cache.update({"key": key, "raw": snap})
        return snap
    lib = _load_readest_lib()
    try:
        import sqlite3
        uri = "file:" + READEST_DB.replace("\\", "/") + "?mode=ro"
        con = sqlite3.connect(uri, uri=True, timeout=2)
        try:
            rows = con.execute(
                "select b.md5, b.title, b.authors, b.pages, p.page, p.start_time, p.duration, p.total_pages "
                "from book b left join page_stat_data p on p.id_book = b.id "
                "order by b.id, p.start_time").fetchall()
        finally:
            con.close()
    except Exception:
        # Readest 没装 / 库被独占: 静默当没有, 书单照常打开
        _rd_cache.update({"key": key, "raw": snap})
        return snap

    books = {}
    for md5, title, authors, pages, page, start, dur, tp in rows:
        md5 = str(md5 or "")
        rec = books.setdefault(md5, {"md5": md5, "db_title": str(title or ""),
                                     "db_author": str(authors or ""),
                                     "pages": int(pages or 0), "events": []})
        if page is None or start is None:
            continue
        rec["events"].append((int(page), int(start), int(dur or 0), int(tp or 0)))

    out = []
    for md5, rec in books.items():
        ev = rec["events"]
        if not ev:
            continue
        lb = lib.get(md5) or {}
        days, total, max_page, total_pages = {}, 0, 0, rec["pages"]
        first_finish = None
        for page, start, dur, tp in ev:
            d = _day(start)
            days[d] = days.get(d, 0) + dur
            total += dur
            if page > max_page:
                max_page = page
            if tp:
                total_pages = max(total_pages, tp)
            if tp and page >= tp and first_finish is None:
                first_finish = d
        qual = sorted(d for d, s in days.items() if s >= READEST_DAY_MIN)
        if not qual:                       # 只是点开瞄了一眼, 不算读过
            continue
        done = ""
        if str(lb.get("readingStatus") or "") == "finished":
            ms = lb.get("readingStatusUpdatedAt")
            if isinstance(ms, (int, float)) and ms > 0:
                done = _day(ms / 1000.0)
        if first_finish and (not done or first_finish < done):
            done = first_finish
        out.append({
            "md5": md5,
            "title": str(lb.get("title") or rec["db_title"] or ""),
            "author": str(lb.get("author") or rec["db_author"] or ""),
            "db_title": rec["db_title"],
            "status": str(lb.get("readingStatus") or ""),
            "进度": [max_page, total_pages or max_page],
            "开始日": qual[0],
            "最近日": qual[-1],
            "读完日": done,
            "天数": len(qual),
            "秒": total,
            "每日": days,
        })
    out.sort(key=lambda x: (-x["秒"], x["title"]))
    snap = {"available": True, "computed_at": int(time.time()), "books": out}
    _rd_cache.update({"key": key, "raw": snap})
    return snap


def _rd_compact(b):
    return {k: b[k] for k in ("开始日", "最近日", "读完日", "天数", "秒", "进度", "每日")}


def _rd_merge(a, b):
    """同一本书在 Readest 里有两份记录(改过文件/换了副本)时, 合成一份。"""
    daily = dict(a["每日"])
    for d, s in b["每日"].items():
        daily[d] = daily.get(d, 0) + s
    qual = [d for d, s in daily.items() if s >= READEST_DAY_MIN]
    starts = [x for x in (a["开始日"], b["开始日"]) if x]
    lasts = [x for x in (a["最近日"], b["最近日"]) if x]
    dones = [x for x in (a["读完日"], b["读完日"]) if x]
    return {"开始日": min(starts) if starts else "",
            "最近日": max(lasts) if lasts else "",
            "读完日": min(dones) if dones else "",
            "天数": len(set(qual)),
            "秒": a["秒"] + b["秒"],
            "进度": b["进度"] if b["进度"][0] > a["进度"][0] else a["进度"],
            "每日": daily}


def _rd_totals(books, days=None):
    days = days or {}
    qual = {d: s for d, s in days.items() if s >= READEST_DAY_MIN}
    month = time.strftime("%Y-%m")
    d0 = datetime.date.today()
    if d0.isoformat() not in qual:
        d0 -= datetime.timedelta(days=1)     # 今天还没读, 连续天数从昨天算
    streak = 0
    while d0.isoformat() in qual:
        streak += 1
        d0 -= datetime.timedelta(days=1)
    return {"秒": sum(int(st.get("秒") or 0) for st in books.values()),
            "天": len(qual),
            "本": len(books),
            "读完": sum(1 for st in books.values() if st.get("读完日")),
            "本月天": sum(1 for d in qual if d.startswith(month)),
            "本月秒": sum(s for d, s in days.items() if d.startswith(month)),
            "连续天": streak,
            "最近日": max(qual) if qual else ""}


def readest_stats(books):
    """原始聚合 + 书单匹配 → 书单视图要的那份统计(完整版)。"""
    raw = readest_raw()
    if not raw["available"]:
        return {"available": False, "computed_at": raw["computed_at"],
                "totals": _rd_totals({}), "days": {}, "books": {}, "unmatched": []}
    by_norm = {}
    for b in books:
        by_norm.setdefault(norm_title(b["书名"]), []).append(b)
    out, unmatched = {}, []
    for rb in raw["books"]:
        cands = by_norm.get(norm_title(rb["title"])) or by_norm.get(norm_title(rb["db_title"])) or []
        hit = None
        if len(cands) == 1:
            hit = cands[0]
        elif cands:
            a = norm_title(rb["author"])
            for b in cands:
                ba = norm_title(b["作者"])
                if a and ba and (a in ba or ba in a):
                    hit = b
                    break
            if hit is None:
                hit = cands[0]
        if hit is None:
            unmatched.append({"title": rb["title"], "author": rb["author"],
                              "开始日": rb["开始日"], "最近日": rb["最近日"],
                              "读完日": rb["读完日"], "天数": rb["天数"],
                              "秒": rb["秒"], "进度": rb["进度"]})
            continue
        name = hit["书名"]
        out[name] = _rd_merge(out[name], rb) if name in out else _rd_compact(rb)
    days = {}
    for st in out.values():
        for d, s in st["每日"].items():
            days[d] = days.get(d, 0) + s
    return {"available": True, "computed_at": raw["computed_at"],
            "totals": _rd_totals(out, days), "days": days,
            "books": out, "unmatched": unmatched}


def _load_auto():
    try:
        with io.open(AUTO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save_auto(d):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = AUTO_FILE + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="\n") as f:
        json.dump(d, f, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, AUTO_FILE)


def _writeback(st, books, stats):
    """只写「开始日 / 读完日」两个字段; 手改过的一律不覆盖。返回改动条数。"""
    auto = _load_auto()
    by_name = {b["书名"]: b for b in books}
    changed, auto_dirty = 0, False
    for name, s in stats["books"].items():
        b = by_name.get(name)
        rec = st["items"].get(name)
        if not b or b.get("判定") == "踢掉" or (rec and rec.get("判定") == "踢掉"):
            continue
        vals = {"开始日": s.get("开始日") or "", "读完日": s.get("读完日") or ""}
        if not any(vals.values()):
            continue
        if rec is None:
            rec = blank_rec()
            rec["序号"] = str(b.get("id") or "")
            rec["作者"] = str(b.get("作者") or "").strip()
            rec["加入日期"] = str(b.get("加入日期") or "").strip()
            st["items"][name] = rec
        old = auto.get(name) if isinstance(auto.get(name), dict) else {}
        manual = set()
        raw_manual = old.get("_manual")
        if isinstance(raw_manual, list):
            manual = {x for x in raw_manual if x in ("开始日", "读完日")}
        for f in ("开始日", "读完日"):
            val = vals[f]
            cur = str(rec.get(f) or "")
            was = str(old.get(f) or "")
            # 见过非自动值就先记成人工值: 以后即使被清空, 也不再自动填回。
            if (cur and cur != was) or (not cur and was):
                if f not in manual:
                    manual.add(f)
                    auto_dirty = True
            if f in manual:
                continue
            if val and cur != val and ((not cur and not was) or (was and cur == was)):
                rec[f] = val                 # 从没写过, 或还是上次自动写的值 → 更新
                old[f] = val
                changed += 1
            elif not val and cur and was and cur == was:
                rec[f] = ""                  # 自动写过的读完日, 现在数据说没读完
                old.pop(f, None)
                changed += 1
        if manual:
            old["_manual"] = sorted(manual)
        else:
            old.pop("_manual", None)
        if old:
            if auto.get(name) != old:
                auto[name] = old
                auto_dirty = True
        elif name in auto:
            auto.pop(name, None)
            auto_dirty = True
    if auto_dirty or changed:
        try:
            _save_auto(auto)
        except OSError:
            pass
    return changed


def sync_readest_dates():
    """把 Readest 的开始/读完日补进标记文件。失败绝不影响书单。"""
    if os.environ.get("BOOKLIST_READEST", "1") == "0":
        return 0
    books, _ = load_books()
    stats = readest_stats(books)
    if not stats["available"] or not stats["books"]:
        return 0
    with _lock:
        try:
            st = load_store(quarantine=False)
        except StoreError:
            return 0
        n = _writeback(st, books, stats)
        if n:
            try:
                save_store(st)
            except OSError:
                return 0
    return n


def fmt_secs(sec):
    sec = int(sec or 0)
    if sec <= 0:
        return "0 分"
    h, m = sec // 3600, int(round((sec % 3600) / 60.0))
    if m == 60:
        h, m = h + 1, 0
    if h and m:
        return "%d 小时 %d 分" % (h, m)
    if h:
        return "%d 小时" % h
    return "%d 分" % m


def view():
    """底本 + 标记 合并后的完整视图。

    认领顺序(记录键 → 哪本书):精确书名 → 归一化书名 → 记录里的「序号」。
    三条通路都要过 fp_ok(作者/加入日期 指纹), 对不上就继续往下走, 最后进 orphans。
    改书名、加书名号、改副标题照样认得回来(序号没变时指纹一定相符); 认领不上的显式列出来。
    """
    books, src = load_books()
    st = load_store()
    rn = renames(st)
    books = [dict(b) for b in books]        # 拷贝: 改名只改视图, 不写回 load_books 的缓存
    for b in books:
        if b.get("主题") and b["主题"] != "未分类":
            b["主题"] = disp(b["主题"], rn)
    items = st["items"]
    used, taken = set(), set()

    def give(b, key, rec):
        used.add(key)
        taken.add(b["书名"])
        b.update({k: (copy_log(rec[k]) if k == "阅读记录" else rec[k]) for k in REC_SHOW})
        b["清单"] = merge_lists(b["主题"], [disp(n, rn) for n in (rec.get("清单") or [])])

    for b in books:
        r = items.get(b["书名"])
        if r is not None and not is_empty(r) and fp_ok(b, r):
            give(b, b["书名"], r)

    folded = {}
    for k in items:
        folded.setdefault(norm_title(k), []).append(k)
    for b in books:                            # 书名被改过(加书名号/改标点)
        if b["书名"] in taken:
            continue
        for k in folded.get(norm_title(b["书名"]), []):
            if k not in used and not is_empty(items[k]) and fp_ok(b, items[k]):
                give(b, k, items[k])
                break

    by_id = {b["id"]: b for b in books if b["id"]}
    for k, v in items.items():                 # 书名改动更大 → 按记录里的序号认领
        if k in used or is_empty(v):
            continue
        b = by_id.get(str(v.get("序号") or ""))
        if b is not None and b["书名"] not in taken and fp_ok(b, v):
            give(b, k, v)

    for b in books:
        if b["书名"] not in taken:
            b.update(BLANK_SHOW)
            b["阅读记录"] = []
            b["清单"] = merge_lists(b["主题"], [])
    orphans = []
    for k, v in items.items():
        if k not in used and not is_empty(v):
            o = dict(书名=k, **v)
            o["阅读记录"] = copy_log(v.get("阅读记录"))
            o["清单"] = merge_lists("", [disp(n, rn) for n in (v.get("清单") or [])])
            o["主题"] = "未分类"                   # 不在底本里, 就没有备注可解析
            orphans.append(o)
    stats = readest_stats(books)
    covers = load_cover_index().get("books", {})
    for b in books:
        b["阅读"] = stats["books"].get(b["书名"])
        b["封面"] = cover_brief(covers.get(str(b.get("id") or "")), b)
    return {"books": books, "source": src, "saved_at": st.get("saved_at", ""),
            "orphans": orphans, "verdicts": list(VERDICTS),
            "readest_available": stats["available"],
            "readest_totals": stats["totals"],
            "readest_unmatched": stats["unmatched"]}


def sort_no(s):
    """序号排序键: 数字按数值排, 非数字(手打「甲」)排后面, 不许抛错。"""
    s = str(s or "")
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


def export_rows(v):
    """导出统一口径: 底本里的记录 + 孤儿, 踢掉的书不摆日期(三条路一致)。

    阅读记录不一样 —— 那是「我哪天真的读了」的流水, 属于笔记, 踢掉也留着。
    """
    rows = []
    for b in v["books"]:
        if not (b["判定"] or b.get("评价") or b["感想"].strip() or b["开始日"]
                or b["读完日"] or b.get("阅读记录")):
            continue
        dead = b["判定"] == "踢掉"
        rd = b.get("阅读") or {}
        rows.append({"书名": b["书名"], "作者": b["作者"], "主题": b["主题"],
                     "清单": "、".join(b.get("清单") or []), "序号": b["id"],
                     "判定": b["判定"], "评价": b.get("评价") or "", "感想": b["感想"],
                     "开始日": "" if dead else b["开始日"],
                     "读完日": "" if dead else b["读完日"],
                     "阅读天数": rd.get("天数", 0),
                     "阅读时长": fmt_secs(rd.get("秒", 0)) if rd else "",
                     "阅读记录": copy_log(b.get("阅读记录")),
                     "更新时间": b["更新时间"], "在底本里": True})
    for o in v["orphans"]:
        dead = o["判定"] == "踢掉"
        rows.append({"书名": o["书名"], "作者": o.get("作者") or "", "主题": "",
                     "清单": "、".join(o.get("清单") or []), "序号": o.get("序号") or "",
                     "判定": o["判定"], "评价": o.get("评价") or "", "感想": o["感想"],
                     "开始日": "" if dead else o["开始日"],
                     "读完日": "" if dead else o["读完日"],
                     "阅读天数": 0, "阅读时长": "",
                     "阅读记录": copy_log(o.get("阅读记录")),
                     "更新时间": o["更新时间"], "在底本里": False})
    return rows


def export_md(v):
    rows = export_rows(v)
    inbook = [r for r in rows if r["在底本里"]]
    orph = [r for r in rows if not r["在底本里"]]
    lines = ["# 我的书单 · 标记与感想", ""]
    for r in sorted(inbook, key=lambda x: (x["清单"], sort_no(x["序号"]))):
        meta = [x for x in (r["清单"], r["判定"]) if x]
        if r["评价"]:
            meta.append("评价 " + r["评价"])
        if r["开始日"]:
            meta.append("开始 " + r["开始日"])
        if r["读完日"]:
            meta.append("读完 " + r["读完日"])
        if r.get("阅读天数"):
            meta.append("读了 %d 天" % r["阅读天数"])
        if r.get("阅读时长"):
            meta.append("有效阅读 " + r["阅读时长"])
        lines.append("- **%s**（%s） — %s" % (r["书名"], r["作者"] or "—", " · ".join(meta)))
        for e in r["阅读记录"]:
            line = e["日期"] + ((" " + e["页码"]) if e["页码"] else "")
            if e["备注"]:
                line += "　" + e["备注"]
            lines.append("  - 读过：" + line)
        if r["感想"].strip():
            for ln in r["感想"].strip().splitlines():
                lines.append("  " + ln)
        lines.append("")
    for r in orph:                             # 底本里已经没有本书了, 笔记也不能从导出里消失
        lines.append("- **%s**（底本里已没有这本） — %s" % (r["书名"], r["判定"] or "—"))
        for e in r["阅读记录"]:
            line = e["日期"] + ((" " + e["页码"]) if e["页码"] else "")
            if e["备注"]:
                line += "　" + e["备注"]
            lines.append("  - 读过：" + line)
        if (r["感想"] or "").strip():
            for ln in r["感想"].strip().splitlines():
                lines.append("  " + ln)
        lines.append("")
    if not rows:
        lines.append("（还没有任何标记或感想）")
    return "\n".join(lines) + "\n"


def export_json(v):
    """和 md 同源: 孤儿也在, 被踢掉的不带日期, 键用「序号·书名」(同名书不会互相盖掉)。"""
    data = {}
    for r in export_rows(v):
        base = ("%s·%s" % (r["序号"], r["书名"])) if r["序号"] else ("%s·%s" % ("孤儿", r["书名"]))
        key, i = base, 2
        while key in data:                     # 真遇上同名同号, 不静默合并
            key, i = "%s#%d" % (base, i), i + 1
        data[key] = {k: r[k] for k in ("书名", "清单", "判定", "评价", "感想", "开始日", "读完日",
                                       "阅读天数", "阅读时长", "阅读记录", "更新时间", "在底本里")}
    return data


SNAP = ("判定", "评价", "感想", "开始日", "读完日", "清单", "阅读记录")
ALL_FIELDS = SNAP


def snap(rec):
    """记录里「会被撤销回滚」的那几个字段的快照(列表要拷一份, 不然撤销时被就地改掉)。"""
    out = {}
    for k in SNAP:
        v = rec.get(k)
        if k == "清单":
            out[k] = list(v or [])
        elif k == "阅读记录":
            out[k] = copy_log(v)
        else:
            out[k] = str(v or "")
    return out


class Conflict(Exception):
    """页面上的信息和底本对不上了(409): 不瞎写, 让页面刷新再试。"""


def parse_updates(payload, allow=ALL_FIELDS):
    """校验并规整一份字段改动。任何一处不合法都抛 ValueError → 整单拒绝, 不留半截写入。"""
    for k in SNAP:
        if k in payload and k not in allow:
            raise ValueError("这个接口不收「%s」" % k)
    updates = {}
    if "判定" in payload:
        v = str(payload.get("判定") or "").strip()
        v = VERDICT_LEGACY.get(v, v)
        if v not in VERDICTS:
            raise ValueError("bad 判定: %s" % v)
        updates["判定"] = v
    if "评价" in payload:
        e = str(payload.get("评价") or "").strip()
        if e not in EVALS:
            raise ValueError("bad 评价: %s" % e)
        updates["评价"] = e
    elif updates.get("判定") == "":
        updates["评价"] = ""          # 退回「未标」= 这本书上什么都没记, 评价跟着清掉
    if "感想" in payload:
        txt = payload.get("感想")
        if txt is None:
            txt = ""
        if not isinstance(txt, str):
            raise ValueError("感想 must be string")
        updates["感想"] = txt[:200000]
    for k in ("开始日", "读完日"):
        if k in payload:
            d = str(payload.get(k) or "").strip()
            if d and not date_ok(d):
                raise ValueError("%s 要真实存在的日期: 2026 / 2026-09 / 2026-09-28" % k)
            updates[k] = d
    if "清单" in payload:
        updates["清单"] = clean_lists(payload.get("清单"))
    if "阅读记录" in payload:
        updates["阅读记录"] = clean_readlog(payload.get("阅读记录"), strict=True)
    return updates


def chrono_bad(rec):
    """「读完」早于「开始读」在两边都是完整日期时一定写错了。

    只比完整日期: 粗粒度的 2026-09 没法跟 2026-09-05 比出个准数, 就先放过。
    """
    a, z = rec.get("开始日") or "", rec.get("读完日") or ""
    if len(a) == 10 and len(z) == 10 and z < a:
        return "读完日期(%s)比开始读(%s)还早, 是不是写反了" % (z, a)
    return ""


def resolve_target(books, items, name, sent_no):
    """书名 → (记录键, 底本那一行或 None, 这条记录)。认领规则和界面同源, 对不上就 Conflict。"""
    by_name = {b["书名"]: b for b in books}
    by_id = {b["id"]: b for b in books if b["id"]}
    b = by_name.get(name)
    rec = items.get(name) or blank_rec()
    if b is not None:
        if sent_no and b["id"] and str(sent_no) != b["id"]:
            raise Conflict("这本书在底本里的位置变了(书单表改过?),按 F5 刷新一下再写")
        return name, b, rec
    cand = by_id.get(str(rec.get("序号") or ""))
    if cand is not None and fp_ok(cand, rec):
        return cand["书名"], cand, rec        # 底本改过名 → 顺手把键迁到新名下
    if name in items:
        return name, None, rec                # 本来就是孤儿记录, 就地改
    raise Conflict("底本里找不到「%s」(可能刚改过书名),按 F5 刷新一下再写" % name)


def apply_updates(items, books, name, sent_no, updates):
    """把一份改动落到 store.items 上, 返回 (记录键, 记录, 改前, 改后, 时间戳)。

    调用方负责加锁、以及最后 save_store —— 所以中途抛 Conflict 时文件一个字节都没动。
    """
    key, b, rec = resolve_target(books, items, name, sent_no)
    before = snap(rec)
    if b is not None and b["id"]:
        # 同一本书因为改过名而攒出两个键时: 取有内容的那条当底, 旧键删掉
        for k in [k for k, v in list(items.items())
                  if k != key and str(v.get("序号") or "") == b["id"]]:
            other = items.pop(k)
            if not is_empty(other):
                other.update(updates)
                rec = other
        rec["序号"] = b["id"]
        rec["作者"] = b["作者"]
        rec["加入日期"] = b["加入日期"]
    rec.update(updates)
    if "读完日" in updates:            # 只在这次真动了「读完日」时比: 把书从「读完」切回「在读」
        msg = chrono_bad(rec)         # 会自动补一个今天的开始日, 那不该被这条挡下来
        if msg:
            raise ValueError(msg)
    now = time.strftime("%Y-%m-%d %H:%M:%S")
    rec["更新时间"] = now
    if is_empty(rec):                         # 全空就删掉, 不留垃圾记录
        items.pop(key, None)
        rec = blank_rec()
    else:
        items[key] = rec
    return key, rec, before, snap(rec), now


class Srv(ThreadingHTTPServer):
    daemon_threads = True
    # 默认 request_queue_size=5 在 Windows 上会让突发并发直接被拒(实测 30 并发只进 6 个)
    request_queue_size = 64


class H(BaseHTTPRequestHandler):
    server_version = "booklist/2.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, val in (extra or {}).items():
            self.send_header(k, val)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _file(self, name, ctype, download=None):
        fp = os.path.join(ROOT, name)
        if not os.path.exists(fp):
            self._send(404, json.dumps({"error": "missing " + name}))
            return
        with open(fp, "rb") as f:
            data = f.read()
        extra = {"Content-Disposition": 'attachment; filename="%s"' % download} if download else None
        self._send(200, data, ctype, extra)

    def _cover_asset(self, rel):
        """只允许读 data/covers 下的 WebP, 不接受 ../ 或任意扩展名。"""
        rel = str(rel or "").replace("\\", "/")
        if not rel.startswith("/covers/") or ".." in rel or not rel.endswith(".webp"):
            self._send(404, json.dumps({"error": "missing cover"}))
            return
        name = rel[len("/covers/"):]
        fp = os.path.abspath(os.path.join(COVER_DIR, name))
        base = os.path.abspath(COVER_DIR) + os.sep
        if not fp.startswith(base) or not os.path.isfile(fp):
            self._send(404, json.dumps({"error": "missing cover"}))
            return
        with open(fp, "rb") as f:
            data = f.read()
        self._send(200, data, "image/webp", {"Cache-Control": "public, max-age=86400"})

    def do_GET(self):
        p = self.path.split("?")[0]
        if p in ("/", "/index.html"):
            self._file("index.html", "text/html; charset=utf-8")
        elif p == "/favicon.ico":
            self._file("app.ico", "image/x-icon")
        elif p in ("/rail-bookshelf.webp", "/paper-grain.webp", "/detail-paper.webp"):
            self._file(p[1:], "image/webp")
        elif p.startswith("/covers/"):
            self._cover_asset(p)
        elif p == "/api/books":
            try:
                try:
                    sync_readest_dates()      # 只补开始日/读完日; 出错也照常给书单
                except Exception:
                    pass
                self._send(200, json.dumps(view(), ensure_ascii=False))
            except StoreError as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
        elif p == "/api/readest":
            try:
                books, _src = load_books()
                self._send(200, json.dumps(readest_stats(books), ensure_ascii=False))
            except Exception as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
        elif p == "/api/store":
            try:
                self._send(200, json.dumps(load_store(), ensure_ascii=False, indent=1))
            except StoreError as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
        elif p == "/api/cover/candidates":
            qs = parse_qs(urlparse(self.path).query)
            book_id = str((qs.get("id") or [""])[0]).strip()
            title = str((qs.get("title") or [""])[0]).strip()
            idx = load_cover_index()
            entry = idx.get("books", {}).get(book_id) if book_id else None
            if not entry and title:
                books, _src = load_books()
                hit = next((b for b in books if b.get("书名") == title), None)
                if hit:
                    book_id = str(hit.get("id") or "")
                    entry = idx.get("books", {}).get(book_id)
            if not isinstance(entry, dict):
                self._send(404, json.dumps({"error": "没有这本书的封面记录"}, ensure_ascii=False))
                return
            self._send(200, json.dumps({
                "id": book_id,
                "state": entry.get("state") or "placeholder",
                "selected": entry.get("selected") or {},
                "candidates": [c for c in (entry.get("candidates") or [])
                               if isinstance(c, dict) and not c.get("ignored")],
            }, ensure_ascii=False))
        elif p == "/api/export.md":
            try:
                self._send(200, export_md(view()), "text/markdown; charset=utf-8",
                           {"Content-Disposition": 'attachment; filename="booklist.md"'})
            except StoreError as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
        elif p == "/api/export.json":
            try:
                data = export_json(view())
            except StoreError as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
                return
            self._send(200, json.dumps(data, ensure_ascii=False, indent=1),
                       "application/json; charset=utf-8",
                       {"Content-Disposition": 'attachment; filename="booklist.json"'})
        else:
            self._send(404, json.dumps({"error": "not found"}))

    def do_POST(self):
        p = self.path.split("?")[0]
        if p not in ("/api/save", "/api/relink", "/api/bulk", "/api/rename", "/api/delete",
                     "/api/cover/select"):
            self._send(404, json.dumps({"error": "not found"}))
            return
        n = 0
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n > MAX_BODY:
                # 必须把 body 读干净再回, 否则 Windows 直接 RST, 客户端连状态码都拿不到
                left = min(n, MAX_BODY * 4)
                while left > 0:
                    chunk = self.rfile.read(min(65536, left))
                    if not chunk:
                        break
                    left -= len(chunk)
                self.close_connection = True
                self._send(413, json.dumps({"error": "body too large"}, ensure_ascii=False))
                return
            payload = json.loads(self.rfile.read(n).decode("utf-8") or "{}")
            if not isinstance(payload, dict):
                raise ValueError("payload must be object")
        except Exception as e:                     # 走到这里 body 已经读完了(或本来就没有)
            self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
            return

        if p == "/api/cover/select":
            book_id = str(payload.get("id") or "").strip()
            action = str(payload.get("action") or "").strip()
            pick = payload.get("candidate")
            if not book_id or action not in ("adopt", "ignore", "restore"):
                self._send(400, json.dumps({"error": "封面操作要带 id 和 adopt/ignore/restore"},
                                           ensure_ascii=False))
                return
            with _lock:
                idx = load_cover_index()
                entry = idx.get("books", {}).get(book_id)
                if not isinstance(entry, dict):
                    self._send(404, json.dumps({"error": "没有这本书的封面记录"}, ensure_ascii=False))
                    return
                candidates = entry.get("candidates") if isinstance(entry.get("candidates"), list) else []
                if action == "adopt":
                    key = str(pick or "").strip()
                    hit = next((c for c in candidates if isinstance(c, dict)
                                and str(c.get("key") or c.get("remote") or "") == key), None)
                    if not hit or not hit.get("thumb") or not hit.get("full"):
                        self._send(400, json.dumps({"error": "候选封面不存在或没有缓存"},
                                                   ensure_ascii=False))
                        return
                    entry["selected"] = {
                        "thumb": hit["thumb"], "full": hit["full"],
                        "source": hit.get("source") or "manual",
                        "confidence": hit.get("confidence") or 1,
                        "isbn": hit.get("isbn") or "",
                        "width": hit.get("width") or 0,
                        "height": hit.get("height") or 0,
                        "remote": hit.get("remote") or "",
                        "title": hit.get("title") or "",
                        "author": hit.get("author") or "",
                        "year": hit.get("year") or "",
                    }
                    entry["state"] = "ready"
                    entry["manual"] = {
                        "action": "adopt",
                        "candidate": key,
                        "at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    }
                elif action == "ignore":
                    key = str(pick or "").strip()
                    found = False
                    for c in candidates:
                        if isinstance(c, dict) and str(c.get("key") or c.get("remote") or "") == key:
                            c["ignored"] = True
                            found = True
                    if not found:
                        self._send(400, json.dumps({"error": "候选封面不存在"},
                                                   ensure_ascii=False))
                        return
                    ignored = set(str(x) for x in (entry.get("ignored_keys") or []))
                    ignored.add(key)
                    entry["ignored_keys"] = sorted(ignored)
                else:
                    entry["state"] = "placeholder"
                    entry["selected"] = {}
                    entry["manual"] = {
                        "action": "restore",
                        "candidate": "",
                        "at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    }
                entry["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                idx["generated_at"] = entry["updated_at"]
                try:
                    save_cover_index(idx)
                except OSError as e:
                    self._send(500, json.dumps({"error": "写不进封面索引: %s" % e},
                                               ensure_ascii=False))
                    return
            self._send(200, json.dumps({"ok": True, "封面": entry}, ensure_ascii=False))
            return

        if p == "/api/bulk":
            # 批量打状态。顶层字段对每一条都生效; 某一条自带字段就以它自己的为准
            # (N 本书一起改判定时, 只有几本要带不一样的日期)。全部先在内存里改, 有一本
            # 对不上就整单不收 —— 文件一个字节都没动(要么全成, 要么一个都不动)。
            items = payload.get("items")
            if not isinstance(items, list) or not items:
                self._send(400, json.dumps({"error": "items 要是一串 {书名, 序号?}"}, ensure_ascii=False))
                return
            if len(items) > 500:
                self._send(400, json.dumps({"error": "一次最多 500 本"}, ensure_ascii=False))
                return
            targets = []
            try:
                # 顶层字段(含 items 之外的杂项)也走一遍校验: 批量里夹带「感想」要当场拒掉,
                # 不能默默忽略 —— 那种「我以为整批写进去了」的坑最难查
                shared = parse_updates(payload, allow=("判定", "开始日", "读完日"))
                for it in items:
                    if not isinstance(it, dict):
                        raise ValueError("items 里有一条不是对象")
                    nm = str(it.get("书名") or "").strip()
                    if not nm:
                        raise ValueError("items 里有一条没有书名")
                    for k in ("感想", "清单", "阅读记录"):
                        if k in it:
                            raise ValueError("批量不收「%s」(「%s」这条要单独写)" % (k, nm))
                    patch = dict(shared)
                    patch.update(parse_updates({k: it[k] for k in ("判定", "开始日", "读完日") if k in it},
                                               allow=("判定", "开始日", "读完日")))
                    if not patch:
                        raise ValueError("「%s」这一条没有要改的字段" % nm)
                    targets.append((nm, str(it.get("序号") or ""), patch))
            except ValueError as e:
                self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
                return

            with _lock:
                try:
                    st = load_store()
                except StoreError as e:
                    self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
                    return
                books, _s = load_books()
                changes, seen, now = [], set(), ""
                try:
                    for nm, no, patch in targets:
                        if nm in seen:
                            continue
                        seen.add(nm)
                        key, _rec, before, after, now = apply_updates(st["items"], books, nm, no, patch)
                        changes.append({"书名": key, "前": before, "后": after})
                except Conflict as e:
                    self._send(409, json.dumps({"error": str(e)}, ensure_ascii=False))
                    return
                except ValueError as e:        # 字段本身不合法(如读完日早于开始读): 文件没动
                    self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
                    return
                st["saved_at"] = now
                try:
                    save_store(st)
                except OSError as e:
                    self._send(500, json.dumps({"error": "写不进标记文件: %s" % e}, ensure_ascii=False))
                    return
            self._send(200, json.dumps({"ok": True, "改动": changes}, ensure_ascii=False))
            return

        if p == "/api/rename":
            # 给「清单 / 主题」改个名字。只动视图层的映射表, 底本 xlsx 一个字节不碰:
            # 底本备注里的「主题：X」是只读的, 改不了; 所以记住「X 以后显示成 Y」。
            frm = str(payload.get("从") or "").strip()
            to = str(payload.get("到") or "").strip()
            if not frm or not to:
                self._send(400, json.dumps({"error": "改名要写清「从」和「到」"}, ensure_ascii=False))
                return
            if len(to) > LIST_NAME_MAX:
                self._send(400, json.dumps({"error": "名字最多 %d 个字" % LIST_NAME_MAX}, ensure_ascii=False))
                return
            if any(ch in to for ch in "\r\n"):
                self._send(400, json.dumps({"error": "名字不能换行"}, ensure_ascii=False))
                return
            if frm == to:
                self._send(200, json.dumps({"ok": True, "从": frm, "到": to}, ensure_ascii=False))
                return
            with _lock:
                try:
                    st = load_store()
                except StoreError as e:
                    self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
                    return
                rn = renames(st)
                canon = set()
                books, _s = load_books()
                for b in books:
                    if b["主题"] and b["主题"] != "未分类":
                        canon.add(b["主题"])
                for v in st["items"].values():
                    for n in (v.get("清单") or []):
                        canon.add(n)
                # 这里不能再 canon.add(frm): 加了它, 「源清单根本不存在」也永远
                # 能凑出一个 keys, 于是改名会「成功」把一个不存在的清单写进映射表。
                keys = [k for k in canon if disp(k, rn) == frm]
                if not keys:
                    self._send(400, json.dumps({"error": "没有叫「%s」的清单了, 按 F5 刷新一下" % frm},
                                               ensure_ascii=False))
                    return
                if to in {disp(k, rn) for k in canon if disp(k, rn) != frm}:
                    self._send(400, json.dumps({"error": "已经有一个叫「%s」的清单了" % to}, ensure_ascii=False))
                    return
                mp = st.get("重命名")
                if not isinstance(mp, dict):
                    mp = {}
                    st["重命名"] = mp
                for k in keys:
                    mp[k] = to
                st["saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                try:
                    save_store(st)
                except OSError as e:
                    self._send(500, json.dumps({"error": "写不进标记文件: %s" % e}, ensure_ascii=False))
                    return
            self._send(200, json.dumps({"ok": True, "从": frm, "到": to}, ensure_ascii=False))
            return

        if p == "/api/delete":
            # 从书单里删掉一本书: 底本那一行没了 + 判定/感想/日期/读过哪天全没。
            # 跟「踢掉」不是一回事: 踢掉是「我不读它了」还留着理由和那一行。
            # 顺序不能反 —— 先删底本, 底本删不掉(被 Excel 占着)就整单放弃;
            # 只删标记不删底本, 书还挂在列表里, 就是原来那个「删了跟没删一样」。
            nm = str(payload.get("书名") or "").strip()
            if not nm:
                self._send(400, json.dumps({"error": "missing 书名"}, ensure_ascii=False))
                return
            no = str(payload.get("序号") or "").strip()
            with _lock:
                try:
                    st = load_store()
                except StoreError as e:
                    self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
                    return
                try:
                    gone_base, bak = base_delete([nm], [no])
                except PermissionError:
                    self._send(409, json.dumps(
                        {"error": "底本正被 Excel/WPS 打开, 关掉表格再删"}, ensure_ascii=False))
                    return
                except (OSError, ValueError) as e:
                    self._send(500, json.dumps({"error": "底本删不掉: %s" % e}, ensure_ascii=False))
                    return
                items = st["items"]
                gone = []
                for k in list(items):
                    v = items[k]
                    if k == nm or (no and str(v.get("序号") or "") == no) or norm_title(k) == norm_title(nm):
                        gone.append(k)
                for k in gone:
                    items.pop(k, None)
                if gone_base:
                    books, _src = load_books()
                    resync_ids(st, books)
                st["saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                try:
                    save_store(st)
                except OSError as e:
                    self._send(500, json.dumps({"error": "写不进标记文件: %s" % e}, ensure_ascii=False))
                    return
            self._send(200, json.dumps({"ok": True, "删掉": gone, "底本删掉": gone_base, "备份": bak},
                                       ensure_ascii=False))
            return

        name = str(payload.get("书名") or "").strip()
        if not name:
            self._send(400, json.dumps({"error": "missing 书名"}, ensure_ascii=False))
            return

        if p == "/api/relink":                 # 把孤儿记录挂到底本里某本书上
            to = str(payload.get("到") or "").strip()
            if not to:
                self._send(400, json.dumps({"error": "missing 到(要挂到哪本书)"}, ensure_ascii=False))
                return
            with _lock:
                try:
                    st = load_store()
                except StoreError as e:
                    self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
                    return
                books, _s = load_books()
                target = next((b for b in books if b["书名"] == to), None)
                rec = st["items"].get(name)
                if target is None:
                    self._send(400, json.dumps({"error": "底本里没有「%s」" % to}, ensure_ascii=False))
                    return
                if rec is None:
                    self._send(400, json.dumps({"error": "没有「%s」这条记录" % name}, ensure_ascii=False))
                    return
                rec["序号"] = target["id"]
                rec["作者"] = target["作者"]
                rec["加入日期"] = target["加入日期"]
                rec["更新时间"] = time.strftime("%Y-%m-%d %H:%M:%S")
                st["items"].pop(name, None)     # 旧键让位, 免得同一本书两个键
                st["items"][to] = rec
                st["saved_at"] = rec["更新时间"]
                try:
                    save_store(st)
                except OSError as e:
                    self._send(500, json.dumps({"error": "写不进标记文件: %s" % e}, ensure_ascii=False))
                    return
            self._send(200, json.dumps({"ok": True, "记录": out_rec(books, rec, renames(st)), "书名": to},
                                       ensure_ascii=False))
            return

        # 先校验再落盘: 任何一处不合法都整单拒绝, 不留半截写入
        try:
            updates = parse_updates(payload)
        except ValueError as e:
            self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
            return

        with _lock:                            # 读-改-写整段加锁, 多标签页也不会互相覆盖
            try:
                st = load_store()
            except StoreError as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
                return
            books, _s = load_books()
            try:
                key, rec, _before, _after, now = apply_updates(
                    st["items"], books, name, payload.get("序号"), updates)
            except Conflict as e:
                self._send(409, json.dumps({"error": str(e)}, ensure_ascii=False))
                return
            except ValueError as e:            # 字段本身不合法(如读完日早于开始读): 文件没动
                self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
                return
            st["saved_at"] = now
            try:
                save_store(st)
            except OSError as e:               # store 被别的程序占着时, 至少让客户端知道
                self._send(500, json.dumps({"error": "写不进标记文件: %s" % e}, ensure_ascii=False))
                return
        self._send(200, json.dumps({"ok": True, "记录": out_rec(books, rec, renames(st)), "书名": key},
                                   ensure_ascii=False))


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    books, src = load_books()
    stamp = {"pid": os.getpid(), "started_ts": time.time(),
             "started": time.strftime("%Y-%m-%d %H:%M:%S"),
             "src": {f: os.path.getmtime(os.path.join(ROOT, f))
                     for f in ("server.py", "index.html") if os.path.exists(os.path.join(ROOT, f))}}
    try:
        with io.open(os.path.join(DATA_DIR, ".server-stamp.json"), "w", encoding="utf-8") as f:
            json.dump(stamp, f)
    except OSError:
        pass
    srv = Srv(("127.0.0.1", PORT), H)
    print("booklist v2 http://127.0.0.1:%d books=%d source=%s data=%s"
          % (PORT, len(books), src, STORE))
    srv.serve_forever()


if __name__ == "__main__":
    main()
