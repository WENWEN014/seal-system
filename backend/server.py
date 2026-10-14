# -*- coding: utf-8 -*-
"""
证照印章管理系统 - 后端 v6
新增：管理员鉴权、柜子布局（4排×3位）保存、入库（外部回柜 / 新章入柜）
数据目录：环境变量 SEAL_DATA_DIR，默认 E:\证照印章数据
"""
import os, sys, csv, time, shutil, threading, secrets, re, io
from datetime import datetime, date, timedelta
from pathlib import Path
from urllib.parse import quote

BASE_DIR = Path(os.environ.get("SEAL_DATA_DIR", r"E:\证照印章数据"))
CSV_ITEMS = BASE_DIR / "items.csv"
CSV_LOCS = BASE_DIR / "locations.csv"
CSV_MOVS = BASE_DIR / "movements.csv"

if not CSV_ITEMS.exists():
    sys.stderr.write(f"[FATAL] 找不到 {CSV_ITEMS}\n")
    sys.exit(1)

try:
    from flask import Flask, request, jsonify, send_file, Response
    from flask_cors import CORS
except ImportError:
    sys.stderr.write("[FATAL] 需要安装 flask flask-cors\n")
    sys.exit(1)

app = Flask(__name__, static_folder=None)
CORS(app)
LOCK = threading.Lock()

# ---------- 管理员（界面级权限，内网场景）----------
ADMIN_USER = "二维"
ADMIN_PWD = "JMYZZ"
TOKEN_DAYS = 7
_tokens = {}   # token -> 过期时间戳

# ---------- 柜子固定布局 ----------
CABINETS = ["柜一", "柜二", "柜三", "柜四", "柜五", "柜六"]
ROWS_PER_CAB = 4
COLS_PER_ROW = 3

ITEMS_HEADERS = ["item_id", "公司名称", "家族", "法人", "公司类型", "注册时间",
                 "资料类型", "资料类别", "当前位置ID", "当前位置", "当前位置类型",
                 "状态", "持有人", "出库时间", "更新时间", "备注"]
LOCS_HEADERS = ["id", "名称", "类型", "柜子", "排", "位", "关联家族", "负责人", "备注"]
MOVS_HEADERS = ["事件ID", "时间", "item_id", "公司名称", "资料类型", "动作",
                "起点ID", "起点名称", "起点类型", "终点ID", "终点名称", "终点类型",
                "操作人", "事由", "备注"]
SEAL_TYPES = ["公章", "法人章", "财务章", "合同章", "发票章"]
STAT_IN, STAT_OUT, STAT_OFF = "在库", "在外部", "已缴销"

# ---------- 工具 ----------
def now_str(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def today_str(): return date.today().strftime("%Y-%m-%d")

# ---------- 中文数字自然排序（柜一/柜二…柜十，筐01/筐02…）----------
_CN_MAP = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9}
_CN_RE = re.compile(r"[零一二三四五六七八九十]{1,3}")

def cn_to_int(s):
    """支持 一~九十九 的中文数字；无法识别返回 None"""
    if len(s) == 1:
        return _CN_MAP.get(s)
    if s.endswith("十") and s[0] in _CN_MAP:          # 二十 / 三十
        return _CN_MAP[s[0]] * 10
    if s.startswith("十") and s[1] in _CN_MAP:        # 十一 / 十五
        return 10 + _CN_MAP[s[1]]
    if len(s) == 3 and s[1] == "十" and s[0] in _CN_MAP and s[2] in _CN_MAP:  # 二十三
        return _CN_MAP[s[0]] * 10 + _CN_MAP[s[2]]
    return None

def _zcmp_norm(s):
    """把名称里的中文数字换成 4 位阿拉伯数字，用于自然排序"""
    def rep(m):
        v = cn_to_int(m.group())
        return f"{v:04d}" if v is not None else m.group()
    return _CN_RE.sub(rep, s or "")

def _natkey(s):
    return [(int(t) if t.isdigit() else -1, t)
            for t in re.split(r"(\d+)", _zcmp_norm(s)) if t]

def read_csv(path, headers):
    if not path.exists(): return []
    for _ in range(5):
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f))
            return [{h: (r.get(h) or "").strip() for h in headers} for r in rows]
        except (PermissionError, OSError):
            time.sleep(0.25)
    raise PermissionError(f"无法读取 {path}")

def write_csv(path, headers, rows):
    bak = path.with_suffix(path.suffix + ".bak")
    if path.exists():
        try: shutil.copy2(path, bak)
        except: pass
    last_err = None
    for _ in range(5):
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
                w.writeheader()
                for r in rows: w.writerow({h: r.get(h, "") for h in headers})
            return
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(0.25)
    try:
        with open(Path(__file__).parent / "write_error.log", "a", encoding="utf-8") as f:
            f.write(f"{now_str()} {Path(path).name}: {type(last_err).__name__} "
                    f"winerror={getattr(last_err,'winerror','?')} {last_err}\n")
    except Exception:
        pass
    raise PermissionError(f"无法写入 {path}（请关闭 Excel/WPS 后重试）")

def find_loc(locs, name):
    for l in locs:
        if l["名称"] == name: return l
    return None

def ensure_person(locs, name):
    p = find_loc(locs, name)
    if p: return p["id"]
    n = 1
    while any(l["id"] == f"P{n:03d}" for l in locs): n += 1
    locs.append({"id": f"P{n:03d}", "名称": name, "类型": "人员", "柜子": "",
                 "排": "", "位": "", "关联家族": "", "负责人": "", "备注": "系统自动创建"})
    return f"P{n:03d}"

def ensure_basket(locs, name, cabinet=""):
    b = find_loc(locs, name)
    if b: return b["id"]
    n = 1
    while any(l["id"] == f"L{n:03d}" for l in locs): n += 1
    locs.append({"id": f"L{n:03d}", "名称": name, "类型": "位置", "柜子": cabinet,
                 "排": "", "位": "", "关联家族": "", "负责人": "", "备注": ""})
    return f"L{n:03d}"

def next_item_id(items):
    n = len(items) + 1
    while any(i["item_id"] == f"IT{n:05d}" for i in items): n += 1
    return f"IT{n:05d}"

def append_movement(movs, **kw):
    mid = len(movs) + 1
    while any(m["事件ID"] == f"MV{mid:05d}" for m in movs): mid += 1
    row = {h: "" for h in MOVS_HEADERS}
    row["事件ID"] = f"MV{mid:05d}"
    row["时间"] = kw.get("时间") or now_str()
    for k in ["item_id", "公司名称", "资料类型", "动作", "起点ID", "起点名称",
              "起点类型", "终点ID", "终点名称", "终点类型", "操作人", "事由", "备注"]:
        row[k] = kw.get(k, "")
    movs.append(row)

# ---------- 鉴权 ----------
def issue_token():
    t = secrets.token_hex(16)
    _tokens[t] = (datetime.now() + timedelta(days=TOKEN_DAYS)).timestamp()
    return t

def valid_token(t):
    if not t: return False
    exp = _tokens.get(t)
    if not exp: return False
    if time.time() > exp:
        _tokens.pop(t, None); return False
    return True

@app.route("/api/login", methods=["POST"])
def api_login():
    d = request.get_json() or {}
    u = (d.get("user") or "").strip()
    p = (d.get("pwd") or "").strip()
    if u == ADMIN_USER and p == ADMIN_PWD:
        t = issue_token()
        return jsonify({"ok": True, "token": t, "user": ADMIN_USER,
                        "expireDays": TOKEN_DAYS})
    return jsonify({"ok": False, "msg": "账号或密码不正确"})

@app.route("/api/logout", methods=["POST"])
def api_logout():
    d = request.get_json() or {}
    _tokens.pop((d.get("token") or "").strip(), None)
    return jsonify({"ok": True})

def need_admin():
    """写操作统一鉴权；不通过返回错误响应"""
    d = request.get_json(silent=True) or {}
    t = (d.get("token") or request.headers.get("X-Token") or "").strip()
    if not valid_token(t):
        return jsonify({"ok": False, "msg": "未登录或登录已过期，请重新登录"})
    return None

# ---------- 路由 ----------
@app.route("/")
def index():
    html_path = Path(__file__).parent / "看板.html"
    if not html_path.exists():
        return Response("看板 HTML 不存在", status=500, mimetype="text/plain; charset=utf-8")
    with open(html_path, "rb") as fp:
        return Response(fp.read(), mimetype="text/html; charset=utf-8")

@app.route("/api/health")
def health():
    return jsonify({"ok": True, "dir": str(BASE_DIR), "now": now_str()})

@app.route("/api/load")
def load():
    try:
        items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
        locs = read_csv(CSV_LOCS, LOCS_HEADERS)
        movs = read_csv(CSV_MOVS, MOVS_HEADERS)
    except PermissionError as e:
        return jsonify({"ok": False, "msg": str(e)}), 500
    # 柜子：locations 中类型=柜子 的记录（动态）；旧数据无记录时回退到固定 6 柜
    # 注意：不再回退到固定 6 柜，删掉的柜子不能"复活"
    cab_names = [l["名称"] for l in locs if l["类型"] == "柜子" and l["名称"]]
    cabinets = [{"name": c, "rows": ROWS_PER_CAB, "cols": COLS_PER_ROW}
                for c in sorted(set(cab_names), key=_natkey)]
    families = sorted({i["家族"] for i in items if i["家族"]})
    return jsonify({"ok": True, "items": items, "locations": locs,
                    "movements": movs, "cabinets": cabinets, "families": families})

# ---------- 新增柜子 ----------
@app.route("/api/addcab", methods=["POST"])
def api_addcab():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "柜子名称不能为空"})
    if len(name) > 20: return jsonify({"ok": False, "msg": "名称过长"})
    with LOCK:
        try:
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        if any(l["名称"] == name for l in locs):
            return jsonify({"ok": False, "msg": f"「{name}」已存在（柜子/筐/人员名称不能重复）"})
        n = 1
        while any(l["id"] == f"L{n:03d}" for l in locs): n += 1
        locs.append({"id": f"L{n:03d}", "名称": name, "类型": "柜子", "柜子": "",
                     "排": "", "位": "", "关联家族": "", "负责人": "",
                     "备注": f"{ROWS_PER_CAB}排×{COLS_PER_ROW}位"})
        append_movement(movs, item_id="", 公司名称="", 资料类型="", 动作="新增柜子",
                        起点ID="", 起点名称="", 起点类型="",
                        终点ID=f"L{n:03d}", 终点名称=name, 终点类型="柜子",
                        操作人=operator, 事由=f"新增柜子（{ROWS_PER_CAB}排×{COLS_PER_ROW}位）")
        try:
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "name": name})

# ---------- 通用移动：出库 / 入库 / 调拨 / 缴销 ----------
def do_move(items, locs, movs, item_ids, action, dest_name, dest_type,
            operator, purpose, date_):
    dest_id = ""
    if dest_type == "人员":
        dest_id = ensure_person(locs, dest_name)
    elif dest_name:
        dest_id = ensure_basket(locs, dest_name) if dest_name else ""
    moved = []
    for i in items:
        if i["item_id"] not in item_ids: continue
        src_id, src_name, src_type = i["当前位置ID"], i["当前位置"], (i["当前位置类型"] or "位置")
        i["当前位置ID"] = dest_id
        i["当前位置"] = dest_name
        i["当前位置类型"] = dest_type
        if action == "出库":
            i["状态"] = STAT_OUT; i["持有人"] = dest_name; i["出库时间"] = date_
        elif action in ("入库", "调拨"):
            i["状态"] = STAT_IN; i["持有人"] = ""; i["出库时间"] = ""
        elif action == "缴销":
            i["状态"] = STAT_OFF; i["持有人"] = ""; i["出库时间"] = ""
        i["更新时间"] = now_str()
        if purpose: i["备注"] = purpose
        moved.append(i)
        append_movement(movs, item_id=i["item_id"], 公司名称=i["公司名称"],
                        资料类型=i["资料类型"], 动作=action,
                        起点ID=src_id, 起点名称=src_name, 起点类型=src_type,
                        终点ID=dest_id, 终点名称=dest_name, 终点类型=dest_type,
                        操作人=operator, 事由=purpose)
    return moved

@app.route("/api/move", methods=["POST"])
def api_move():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    item_ids = d.get("item_ids", [])
    action = d.get("action", "")
    dest_name = d.get("dest_name", "")
    dest_type = d.get("dest_type", "位置")
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    purpose = (d.get("purpose") or "").strip()
    date_ = d.get("date", today_str())
    if not item_ids: return jsonify({"ok": False, "msg": "请选择资料"})
    if not action: return jsonify({"ok": False, "msg": "请指定动作"})
    if not dest_name: return jsonify({"ok": False, "msg": "请指定目标位置/人员"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        moved = do_move(items, locs, movs, item_ids, action, dest_name, dest_type,
                        operator, purpose, date_)
        if not moved: return jsonify({"ok": False, "msg": "没有匹配的资料"})
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "moved": len(moved),
                    "detail": f"{len(moved)} 件：{action} → {dest_name}"})

# ---------- 布局保存（拖拽，支持跨柜）----------
@app.route("/api/savelayout", methods=["POST"])
def api_savelayout():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    baskets = d.get("baskets", [])   # [{名称, 柜子, 排, 位}]
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not baskets: return jsonify({"ok": False, "msg": "无布局数据"})
    with LOCK:
        try:
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        changed = []
        for b in baskets:
            nm = (b.get("名称") or "").strip()
            if not nm: continue
            l = find_loc(locs, nm)
            if not l: continue
            old_cab, old_row, old_col = l["柜子"], l["排"], l["位"]
            new_cab = (b.get("柜子") or old_cab).strip()
            new_row = str(b.get("排") or "").strip()
            new_col = str(b.get("位") or "").strip()
            if (old_cab, old_row, old_col) != (new_cab, new_row, new_col):
                l["柜子"], l["排"], l["位"] = new_cab, new_row, new_col
                changed.append(f"{nm}: {old_cab}{old_row}排{old_col}位 → {new_cab}{new_row}排{new_col}位")
        if not changed:
            return jsonify({"ok": True, "changed": 0})
        try:
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
            append_movement(movs, item_id="", 公司名称="", 资料类型="",
                            动作="柜子布局调整", 起点ID="", 起点名称="", 起点类型="",
                            终点ID="", 终点名称="", 终点类型="",
                            操作人=operator, 事由="；".join(changed))
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "changed": len(changed), "detail": changed})

# ---------- 新增公司 ----------
@app.route("/api/newco", methods=["POST"])
def api_newco():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    legal = (d.get("legalPerson") or "").strip()
    family = (d.get("family") or "").strip()
    ctype = d.get("type") or "大陆"
    reg = d.get("regTime", "")
    position = (d.get("position") or "").strip()
    seals = d.get("seals", [])
    want_lic = d.get("license", True)
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "公司名不能为空"})
    if not legal: return jsonify({"ok": False, "msg": "法人不能为空"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        if any(i["公司名称"] == name for i in items):
            return jsonify({"ok": False, "msg": f"公司 {name} 已存在"})
        pos_id = ensure_basket(locs, position) if position else ""
        created = []
        for s in seals:
            items.append({"item_id": next_item_id(items), "公司名称": name, "家族": family,
                          "法人": legal, "公司类型": ctype, "注册时间": reg,
                          "资料类型": s, "资料类别": "章", "当前位置ID": pos_id,
                          "当前位置": position or "待分配", "当前位置类型": "位置",
                          "状态": STAT_IN, "持有人": "", "出库时间": "",
                          "更新时间": now_str(), "备注": ""})
            created.append(s)
            append_movement(movs, item_id=items[-1]["item_id"], 公司名称=name, 资料类型=s,
                            动作="入库", 起点ID="", 起点名称="新注册", 起点类型="",
                            终点ID=pos_id, 终点名称=position or "待分配", 终点类型="位置",
                            操作人=operator, 事由="新注册入库")
        if want_lic:
            items.append({"item_id": next_item_id(items), "公司名称": name, "家族": family,
                          "法人": legal, "公司类型": ctype, "注册时间": reg,
                          "资料类型": "营业执照", "资料类别": "执照", "当前位置ID": "",
                          "当前位置": "待分配", "当前位置类型": "位置", "状态": STAT_IN,
                          "持有人": "", "出库时间": "", "更新时间": now_str(),
                          "备注": "存放柜二，待定位"})
            created.append("营业执照")
            append_movement(movs, item_id=items[-1]["item_id"], 公司名称=name,
                            资料类型="营业执照", 动作="入库", 起点ID="", 起点名称="新注册",
                            起点类型="", 终点ID="", 终点名称="待分配", 终点类型="位置",
                            操作人=operator, 事由="新注册入库")
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "created": created})

# ---------- 编辑 / 增删资料 ----------
@app.route("/api/edititem", methods=["POST"])
def api_edititem():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    item_id = d.get("item_id", "")
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not item_id: return jsonify({"ok": False, "msg": "缺少 item_id"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        t = next((i for i in items if i["item_id"] == item_id), None)
        if not t: return jsonify({"ok": False, "msg": "找不到该资料"})
        changed = []
        for k, cn in [("name", "公司名称"), ("legalPerson", "法人"),
                      ("family", "家族"), ("note", "备注")]:
            v = (d.get(k) or "").strip()
            if v and v != t[cn]:
                changed.append(f"{cn}: {t[cn]} → {v}"); t[cn] = v
        t["更新时间"] = now_str()
        if changed:
            append_movement(movs, item_id=item_id, 公司名称=t["公司名称"],
                            资料类型=t["资料类型"], 动作="编辑", 起点ID="", 起点名称="",
                            起点类型="", 终点ID="", 终点名称="", 终点类型="",
                            操作人=operator, 事由="；".join(changed))
            try:
                write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
                write_csv(CSV_MOVS, MOVS_HEADERS, movs)
            except PermissionError as e:
                return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "changed": changed})

@app.route("/api/additem", methods=["POST"])
def api_additem():
    """给公司新增一枚章/执照（新章入柜场景）"""
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    itype = (d.get("type") or d.get("item") or "").strip()
    position = (d.get("position") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "公司名不能为空"})
    if not itype: return jsonify({"ok": False, "msg": "资料类型不能为空"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        tpl = next((i for i in items if i["公司名称"] == name), None)
        if not tpl: return jsonify({"ok": False, "msg": f"找不到公司：{name}"})
        if any(i["公司名称"] == name and i["资料类型"] == itype for i in items):
            return jsonify({"ok": False, "msg": f"{name} 已有 {itype}"})
        kind = "执照" if itype in ("营业执照", "国外执照") else "章"
        pos = position or (tpl["当前位置"] if kind == "章" else "待分配")
        pos_id = ensure_basket(locs, pos) if pos and pos != "待分配" else ""
        items.append({"item_id": next_item_id(items), "公司名称": name, "家族": tpl["家族"],
                      "法人": tpl["法人"], "公司类型": tpl["公司类型"], "注册时间": tpl["注册时间"],
                      "资料类型": itype, "资料类别": kind, "当前位置ID": pos_id,
                      "当前位置": pos or "待分配", "当前位置类型": "位置", "状态": STAT_IN,
                      "持有人": "", "出库时间": "", "更新时间": now_str(), "备注": ""})
        append_movement(movs, item_id=items[-1]["item_id"], 公司名称=name, 资料类型=itype,
                        动作="入库", 起点ID="", 起点名称="新增", 起点类型="",
                        终点ID=pos_id, 终点名称=pos or "待分配", 终点类型="位置",
                        操作人=operator, 事由="新增资料入库")
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True})

@app.route("/api/delitem", methods=["POST"])
def api_delitem():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    item_id = d.get("item_id", "")
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not item_id: return jsonify({"ok": False, "msg": "缺少 item_id"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        t = next((i for i in items if i["item_id"] == item_id), None)
        if not t: return jsonify({"ok": False, "msg": "找不到该资料"})
        append_movement(movs, item_id=item_id, 公司名称=t["公司名称"], 资料类型=t["资料类型"],
                        动作="删除", 起点ID=t["当前位置ID"], 起点名称=t["当前位置"],
                        起点类型=t["当前位置类型"], 终点ID="", 终点名称="（已删除）",
                        终点类型="", 操作人=operator, 事由="删除资料")
        items = [i for i in items if i["item_id"] != item_id]
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True})

@app.route("/api/delco", methods=["POST"])
def api_delco():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "公司名不能为空"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        n = len([i for i in items if i["公司名称"] == name])
        if not n: return jsonify({"ok": False, "msg": f"找不到公司：{name}"})
        append_movement(movs, item_id="", 公司名称=name, 资料类型="全部", 动作="删除公司",
                        起点ID="", 起点名称="", 起点类型="", 终点ID="", 终点名称="",
                        终点类型="", 操作人=operator, 事由=f"删除公司及其 {n} 件资料")
        items = [i for i in items if i["公司名称"] != name]
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "deleted": n})

@app.route("/api/assignpos", methods=["POST"])
def api_assignpos():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    item_ids = d.get("item_ids", [])
    dest_name = (d.get("dest_name") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not item_ids: return jsonify({"ok": False, "msg": "请选择资料"})
    if not dest_name: return jsonify({"ok": False, "msg": "请选择目标位置"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        moved = do_move(items, locs, movs, item_ids, "调拨", dest_name, "位置",
                        operator, "批量分配位置", today_str())
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "moved": len(moved)})

@app.route("/api/addloc", methods=["POST"])
def api_addloc():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    ltype = d.get("type") or "位置"
    cabinet = (d.get("cabinet") or "").strip()
    family = (d.get("family") or "").strip()
    owner = (d.get("owner") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "名称不能为空"})
    with LOCK:
        try:
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        if find_loc(locs, name):
            return jsonify({"ok": False, "msg": f"{name} 已存在"})
        # 新筐自动放到该柜第一个空位
        row = col = ""
        if ltype == "位置" and cabinet:
            used = {(l["排"], l["位"]) for l in locs
                    if l["柜子"] == cabinet and l["类型"] == "位置" and l["排"]}
            for r in range(1, ROWS_PER_CAB + 1):
                for c in range(1, COLS_PER_ROW + 1):
                    if (str(r), str(c)) not in used:
                        row, col = str(r), str(c); break
                if row: break
        if ltype == "人员":
            ensure_person(locs, name)
        else:
            ensure_basket(locs, name, cabinet)
        nl = find_loc(locs, name)
        if nl:
            nl["排"], nl["位"] = row, col
            if family: nl["关联家族"] = family
            if owner: nl["负责人"] = owner
        append_movement(movs, item_id="", 公司名称="", 资料类型="", 动作="新增位置",
                        起点ID="", 起点名称="", 起点类型="", 终点ID=nl["id"] if nl else "",
                        终点名称=name, 终点类型=ltype, 操作人=operator,
                        事由=f"新增{'人员' if ltype=='人员' else '筐'}")
        try:
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True})

# ---------- 删除筐 / 人员 ----------
@app.route("/api/delloc", methods=["POST"])
def api_delloc():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "名称不能为空"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        l = find_loc(locs, name)
        if not l: return jsonify({"ok": False, "msg": f"找不到「{name}」"})
        if l["类型"] == "柜子":
            return jsonify({"ok": False, "msg": "柜子请用柜子图上的「删除柜」"})
        if l["类型"] == "人员":
            held = [i for i in items if i["状态"] == STAT_OUT and i["持有人"] == name]
            if held:
                return jsonify({"ok": False, "msg": f"「{name}」还持有 {len(held)} 件资料，请先收回再删除"})
            kind = "人员"
        else:
            inside = [i for i in items if i["当前位置"] == name and i["状态"] != STAT_OFF]
            if inside:
                return jsonify({"ok": False, "msg": f"「{name}」里还有 {len(inside)} 件资料，请先调拨或缴销后再删除"})
            kind = "筐"
        locs = [x for x in locs if x["名称"] != name]
        append_movement(movs, item_id="", 公司名称="", 资料类型="", 动作="删除位置",
                        起点ID=l["id"], 起点名称=name, 起点类型=l["类型"],
                        终点ID="", 终点名称="", 终点类型="",
                        操作人=operator, 事由=f"删除{kind}")
        try:
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True})

# ---------- 删除柜子（连带删除柜内空筐）----------
@app.route("/api/delcab", methods=["POST"])
def api_delcab():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    name = (d.get("name") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not name: return jsonify({"ok": False, "msg": "名称不能为空"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        cab = find_loc(locs, name)
        if not cab or cab["类型"] != "柜子":
            return jsonify({"ok": False, "msg": f"找不到柜子「{name}」"})
        baskets = [l for l in locs if l["类型"] == "位置" and l["柜子"] == name]
        busy = [b for b in baskets
                if any(i["当前位置"] == b["名称"] and i["状态"] != STAT_OFF for i in items)]
        if busy:
            names = "、".join(b["名称"] for b in busy[:5])
            return jsonify({"ok": False, "msg": f"柜内还有 {len(busy)} 个非空筐（{names}"
                                                f"{'…' if len(busy) > 5 else ''}），请先把筐拖到别的柜再删除"})
        kill = {name} | {b["名称"] for b in baskets}
        locs = [x for x in locs if x["名称"] not in kill]
        append_movement(movs, item_id="", 公司名称="", 资料类型="", 动作="删除柜子",
                        起点ID=cab["id"], 起点名称=name, 起点类型="柜子",
                        终点ID="", 终点名称="", 终点类型="",
                        操作人=operator, 事由=f"删除柜子（连带 {len(baskets)} 个空筐）")
        try:
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "baskets": len(baskets)})

# ---------- 改名（柜子 / 筐 / 人员）----------
@app.route("/api/renameloc", methods=["POST"])
def api_renameloc():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    old = (d.get("old") or d.get("name") or "").strip()
    new = (d.get("new") or "").strip()
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not old or not new:
        return jsonify({"ok": False, "msg": "原名称和新名称都不能为空"})
    if old == new:
        return jsonify({"ok": False, "msg": "新名称与原名称相同"})
    if len(new) > 20:
        return jsonify({"ok": False, "msg": "名称过长"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        l = find_loc(locs, old)
        if not l:
            return jsonify({"ok": False, "msg": f"找不到「{old}」"})
        if find_loc(locs, new):
            return jsonify({"ok": False, "msg": f"「{new}」已存在（柜子/筐/人员名称不能重复）"})
        kind = l["类型"]
        n_item = 0
        if kind == "柜子":
            for x in locs:
                if x["柜子"] == old: x["柜子"] = new
        elif kind == "位置":
            for i in items:
                if i["当前位置"] == old:
                    i["当前位置"] = new; i["更新时间"] = now_str(); n_item += 1
        elif kind == "人员":
            for i in items:
                if i["持有人"] == old:
                    i["持有人"] = new; i["更新时间"] = now_str(); n_item += 1
                if i["当前位置"] == old:
                    i["当前位置"] = new
        l["名称"] = new
        n_mov = 0
        for m in movs:
            if m["起点名称"] == old: m["起点名称"] = new; n_mov += 1
            if m["终点名称"] == old: m["终点名称"] = new; n_mov += 1
        append_movement(movs, item_id="", 公司名称="", 资料类型="", 动作="改名",
                        起点ID=l["id"], 起点名称=old, 起点类型=kind,
                        终点ID=l["id"], 终点名称=new, 终点类型=kind,
                        操作人=operator,
                        事由=f"{'柜子' if kind=='柜子' else ('人员' if kind=='人员' else '筐')}改名" +
                             (f"；同步 {n_item} 件资料、{n_mov} 条轨迹" if (n_item or n_mov) else ""))
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "items": n_item, "movements": n_mov})

@app.route("/api/batchimp", methods=["POST"])
def api_batchimp():
    guard = need_admin()
    if guard: return guard
    d = request.get_json()
    rows = d.get("rows", [])
    operator = (d.get("operator") or "").strip() or ADMIN_USER
    if not rows: return jsonify({"ok": False, "msg": "无数据"})
    with LOCK:
        try:
            items = read_csv(CSV_ITEMS, ITEMS_HEADERS)
            locs = read_csv(CSV_LOCS, LOCS_HEADERS)
            movs = read_csv(CSV_MOVS, MOVS_HEADERS)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
        created, skipped = [], []
        for r in rows:
            name = (r.get("name") or "").strip()
            if not name: continue
            family = (r.get("family") or "").strip()
            legal = (r.get("legalPerson") or "").strip()
            ctype = r.get("type") or "大陆"
            reg = r.get("regTime", "")
            pos = (r.get("position") or "").strip()
            want = []
            for s in SEAL_TYPES:
                if str(r.get(s, "0")) in ("1", "1.0", "true", "True"):
                    want.append((s, "章"))
            if str(r.get("license", "1")) in ("1", "1.0", "true", "True"):
                want.append(("营业执照", "执照"))
            if not want: continue
            pos_id = ensure_basket(locs, pos) if pos else ""
            for itype, kind in want:
                if any(i["公司名称"] == name and i["资料类型"] == itype for i in items):
                    skipped.append(f"{name}/{itype}"); continue
                tgt = pos if kind == "章" else "待分配"
                tgt_id = pos_id if kind == "章" else ""
                items.append({"item_id": next_item_id(items), "公司名称": name, "家族": family,
                              "法人": legal, "公司类型": ctype, "注册时间": reg,
                              "资料类型": itype, "资料类别": kind, "当前位置ID": tgt_id,
                              "当前位置": tgt or "待分配", "当前位置类型": "位置",
                              "状态": STAT_IN, "持有人": "", "出库时间": "",
                              "更新时间": now_str(), "备注": ""})
                created.append(f"{name}/{itype}")
                append_movement(movs, item_id=items[-1]["item_id"], 公司名称=name,
                                资料类型=itype, 动作="入库", 起点ID="", 起点名称="批量导入",
                                起点类型="", 终点ID=tgt_id, 终点名称=tgt or "待分配",
                                终点类型="位置", 操作人=operator, 事由="批量导入")
        try:
            write_csv(CSV_ITEMS, ITEMS_HEADERS, items)
            write_csv(CSV_LOCS, LOCS_HEADERS, locs)
            write_csv(CSV_MOVS, MOVS_HEADERS, movs)
        except PermissionError as e:
            return jsonify({"ok": False, "msg": str(e)})
    return jsonify({"ok": True, "created": created, "skipped": skipped})

@app.route("/api/export/<name>")
def api_export(name):
    m = {"items": CSV_ITEMS, "locations": CSV_LOCS, "movements": CSV_MOVS}
    p = m.get(name)
    if not p or not p.exists():
        return jsonify({"ok": False, "msg": "文件不存在"})
    return send_file(p, as_attachment=True, download_name=p.name)

# ---------- 盘点表导出 Excel（合并单元格 + 边框）----------
@app.route("/api/invxlsx", methods=["POST"])
def api_invxlsx():
    """前端把算好的矩阵 POST 过来，这里只负责排版成 xlsx。
    入参：{title, header:[...], groups:[{no, co, lead:[柜子,筐...], lines:[[...],[...]]}]}
    列布局：序号 | lead... | 公司 | lines...（header 需与此对应）"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    except ImportError:
        return jsonify({"ok": False, "msg": "服务器缺少 openpyxl，无法导出 Excel"}), 500
    d = request.get_json() or {}
    title = (d.get("title") or "盘点表").strip()
    header = d.get("header") or []
    groups = d.get("groups") or []
    if not header or not groups:
        return jsonify({"ok": False, "msg": "没有可导出的数据"})

    wb = Workbook()
    ws = wb.active
    ws.title = "盘点表"
    thin = Side(style="thin", color="9AA0A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_al = Alignment(horizontal="left", vertical="center")
    fill = PatternFill("solid", fgColor="EFF3F8")

    ncol = len(header)
    # 标题
    ws.cell(row=1, column=1, value=title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(row=1, column=1)
    c.font = Font(name="微软雅黑", size=14, bold=True)
    c.alignment = center
    ws.row_dimensions[1].height = 30
    # 表头
    for j, h in enumerate(header, 1):
        c = ws.cell(row=2, column=j, value=h)
        c.font = Font(name="微软雅黑", size=11, bold=True)
        c.alignment = center
        c.border = border
        c.fill = fill
    ws.row_dimensions[2].height = 26
    # 数据。列布局：1=序号，2..1+lead_n=前置列（柜子/筐），2+lead_n=公司，其后为 lines
    lead_n = 0
    if groups and isinstance(groups[0].get("lead"), list):
        lead_n = len(groups[0]["lead"])
    fixed = 2 + lead_n          # 序号 + 前置列 + 公司（这些列按组合并）
    r = 3
    for g in groups:
        lines = g.get("lines") or []
        if not lines: continue
        span = len(lines)
        lead = g.get("lead") or []
        c = ws.cell(row=r, column=1, value=g.get("no", ""))
        c.alignment = center; c.border = border
        c.font = Font(name="微软雅黑", size=11)
        for j0, v0 in enumerate(lead, 2):
            c = ws.cell(row=r, column=j0, value=v0)
            c.alignment = center; c.border = border
            c.font = Font(name="微软雅黑", size=11)
        c = ws.cell(row=r, column=2 + lead_n, value=g.get("co", ""))
        c.alignment = left_al; c.border = border
        c.font = Font(name="微软雅黑", size=11)
        # 合并列的剩余行也要画边框
        for k in range(1, span):
            for j0 in range(1, fixed + 1):
                ws.cell(row=r + k, column=j0).border = border
        if span > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r + span - 1, end_column=1)
            for j0 in range(2, fixed + 1):
                ws.merge_cells(start_row=r, start_column=j0, end_row=r + span - 1, end_column=j0)
        for li, line in enumerate(lines):
            for j in range(fixed, ncol):
                v = line[j - fixed] if j - fixed < len(line) else ""
                c = ws.cell(row=r + li, column=j + 1, value=v)
                c.alignment = center
                c.border = border
                c.font = Font(name="微软雅黑", size=11)
                if v in ("无", "已缴销"):
                    c.font = Font(name="微软雅黑", size=11, color="9AA0A6")
                elif v and v != "✓" and not v.startswith("✓×"):
                    c.font = Font(name="微软雅黑", size=11, color="BF8700")
            ws.row_dimensions[r + li].height = 22
        r += span
    # 列宽：按表头名匹配
    from openpyxl.utils import get_column_letter
    width_map = {"序号": 6, "柜子": 9, "筐": 17, "公司": 30, "实际核对": 16}
    for j, h in enumerate(header, 1):
        ws.column_dimensions[get_column_letter(j)].width = width_map.get(h, 13)
    ws.freeze_panes = get_column_letter(fixed + 1) + "3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"{title}.xlsx"
    return Response(buf.getvalue(),
                    mimetype="application/vnd.openxmlformats-officedocument"
                             "-spreadsheetml.sheet",
                    headers={"Content-Disposition":
                             "attachment; filename*=UTF-8''" + quote(fname)})

if __name__ == "__main__":
    port = 8000
    print("=== 证照印章系统后端 v6（含鉴权与布局）===")
    print(f"数据目录：{BASE_DIR}")
    print(f"管理员：{ADMIN_USER}（登录有效期 {TOKEN_DAYS} 天）")
    print(f"柜子布局：{len(CABINETS)} 柜 × {ROWS_PER_CAB} 排 × {COLS_PER_ROW} 位")
    print(f"监听：http://0.0.0.0:{port}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)