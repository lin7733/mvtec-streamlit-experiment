
import hashlib
import json
import random
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
from PIL import Image, ImageFile
import streamlit as st

@st.cache_data(show_spinner=False)

def load_image_bytes(image_path_str: str) -> bytes:

    with open(image_path_str, "rb") as f:

        return f.read()

try:
    import gspread
    from gspread.exceptions import WorksheetNotFound
except Exception:
    gspread = None
    WorksheetNotFound = Exception

ImageFile.LOAD_TRUNCATED_IMAGES = True

st.set_page_config(
    page_title="AI辅助质检实验平台",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_FILENAME = "AI_质检实验题库_实验一实验二_多目标采购版_AI正确率75版.xlsx"
WORKBOOK_FALLBACK_FILENAMES = ["AI_质检实验题库_实验一实验二_多目标采购版.xlsx"]
DATASET_ROOT = BASE_DIR / "00_raw_compressed"
RESULTS_DIR_DEFAULT = BASE_DIR / "results"

APP_TITLE = "AI辅助质检实验平台"
FORMAL_TRIALS_PER_EXPERIMENT = 24
PRACTICE_TRIALS_PER_EXPERIMENT = 2
BREAK_AFTER = FORMAL_TRIALS_PER_EXPERIMENT
OUTPUT_XLSX_NAME = "experiment_data.xlsx"
PRODUCT_CATEGORIES = ("bottle", "capsule", "metal_nut")

CONSENT_TEXT = """
**知情同意书**

本实验为本科毕业论文研究项目，研究主题为“AI 辅助决策中解释透明度与依赖行为”。

**实验内容：**
你将完成一系列工业产品判断任务，并在实验结束后填写简短问卷。

**数据使用：**
实验过程中记录的作答结果、反应时间和问卷评分，仅用于学术研究分析。

**自愿参与：**
你可以在任何时候选择退出实验，不会产生任何不良后果。

**实验时长：**
约 20–30 分钟。

请确认你已充分理解以上内容，并自愿参与本实验。
"""

PRODUCT_STANDARDS = {
    "bottle": {
        "name": "瓶子（Bottle）",
        "ok": "瓶身表面完整光滑，无可见污渍、破损或异物附着。",
        "ng_types": [
            "污染（Contamination）：瓶身表面有污渍、附着物或颜色异常区域",
            "大破损（Broken Large）：瓶口、瓶身出现较大缺口或破裂",
            "小破损（Broken Small）：瓶身出现细微裂缝或小缺口",
        ],
    },
    "capsule": {
        "name": "胶囊（Capsule）",
        "ok": "胶囊外壳完整，颜色均匀，无变形、裂缝或内容物渗漏。",
        "ng_types": [
            "裂缝（Crack）：胶囊壳体出现裂纹或断裂",
            "渗漏（Squeeze）：胶囊受挤压变形或内容物外漏",
            "压痕（Poke）：胶囊表面有明显凹陷或戳痕",
            "划痕（Scratch）：胶囊表面有线状划伤痕迹",
        ],
    },
    "metal_nut": {
        "name": "金属螺母（Metal Nut）",
        "ok": "螺母形状规则，颜色均匀，螺纹完整，无弯曲、翻转或划痕。",
        "ng_types": [
            "弯曲（Bent）：螺母出现变形或弯曲",
            "颜色异常（Color）：螺母表面颜色不均匀或有锈斑",
            "翻转（Flip）：螺母放置方向错误或翻面",
            "划痕（Scratch）：螺母表面有明显划伤",
        ],
    },
}


def safe_str(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def stable_hash_int(text: str) -> int:
    return int(hashlib.md5(text.encode("utf-8")).hexdigest(), 16)


def sanitize_for_path(text: str) -> str:
    cleaned = "".join(ch for ch in safe_str(text) if ch.isalnum() or ch in {"_", "-"})
    return cleaned or "unknown"


def build_participant_id(student_id: str, exp_short: str) -> str:
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    return f"{exp_short}_{sanitize_for_path(student_id)}_{ts}"


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [safe_str(c) for c in df.columns]
    for c in df.columns:
        df[c] = df[c].apply(safe_str)
    return df


def find_workbook_path() -> Path:
    workbook_filenames = [WORKBOOK_FILENAME, *WORKBOOK_FALLBACK_FILENAMES]
    candidates = []
    for filename in workbook_filenames:
        candidates.extend(
            [
                BASE_DIR / "05_metadata" / filename,
                BASE_DIR / filename,
            ]
        )
    for p in candidates:
        if p.exists() and p.is_file():
            return p

    metadata_dir = BASE_DIR / "05_metadata"
    if metadata_dir.exists():
        dir_listing = "\n".join(sorted(x.name for x in metadata_dir.iterdir()))
    else:
        dir_listing = "<05_metadata 文件夹不存在>"

    searched = "\n".join(str(p) for p in candidates)
    raise FileNotFoundError(
        "未找到题库文件。\n已搜索路径：\n"
        f"{searched}\n\n"
        "05_metadata 目录内容：\n"
        f"{dir_listing}"
    )


def read_structured_sheet(workbook_path: str, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None)
    header_row_idx = None
    for idx in range(min(5, len(raw))):
        row_values = [safe_str(x) for x in raw.iloc[idx].tolist()]
        if "题号" in row_values or "练习题号" in row_values:
            header_row_idx = idx
            break
    if header_row_idx is None:
        raise ValueError(f"{sheet_name} 未找到表头行。")

    headers = [safe_str(x) for x in raw.iloc[header_row_idx].tolist()]
    data = raw.iloc[header_row_idx + 1 :].copy()
    data.columns = headers
    data = data.dropna(how="all")
    data = normalize_df(data)
    if headers[0] in data.columns:
        data = data[data[headers[0]] != ""]
    return data.reset_index(drop=True)


def to_int(value, default=0):
    try:
        if safe_str(value) == "":
            return default
        return int(float(value))
    except Exception:
        return default


def to_float(value, default=0.0):
    try:
        if safe_str(value) == "":
            return default
        return float(value)
    except Exception:
        return default


def normalize_category(value: str) -> str:
    text = safe_str(value).lower()
    mapping = {
        "bottle": "bottle",
        "capsule": "capsule",
        "metal_nut": "metal_nut",
        "瓶": "bottle",
        "胶囊": "capsule",
        "螺母": "metal_nut",
    }
    for key, val in mapping.items():
        if key in text:
            return val
    return safe_str(value)


def normalize_complexity(value: str) -> str:
    text = safe_str(value).lower()
    if "low" in text or "低" in text:
        return "low"
    if "high" in text or "高" in text:
        return "high"
    return safe_str(value)


def normalize_okng_label(value: str) -> str:
    text = safe_str(value).lower().replace("（", "(").replace("）", ")")
    compact = "".join(text.split())
    if compact in {"0", "ng"}:
        return "NG"
    if compact in {"1", "ok"}:
        return "OK"
    if compact.startswith("ng") or any(k in compact for k in ["不合格", "不正常", "异常", "bad"]):
        return "NG"
    if compact.startswith("ok") or any(k in compact for k in ["无缺陷", "正常", "good", "nodefect"]):
        return "OK"
    if "缺陷" in compact or ("defect" in compact and "nodefect" not in compact):
        return "NG"
    if "合格" in compact:
        return "OK"
    return safe_str(value)


def okng_code(value: str) -> int:
    return 1 if normalize_okng_label(value) == "OK" else 0


def normalize_purchase_label(value: str) -> str:
    text = safe_str(value)
    if text in {"1", "采购", "建议采购", "buy", "BUY"}:
        return "采购"
    if text in {"0", "不采购", "建议不采购", "not_buy", "NO_BUY"}:
        return "不采购"
    return text


def purchase_code(value: str) -> int:
    text = safe_str(value)
    if text in {"1", "采购", "建议采购", "buy", "BUY"}:
        return 1
    if text in {"0", "不采购", "建议不采购", "not_buy", "NO_BUY"}:
        return 0
    return 0


def parse_ai_correct(value: str, fallback=None) -> int:
    text = safe_str(value).lower()
    if any(k in text for k in ["正确", "true", "yes", "是", "1"]):
        return 1
    if any(k in text for k in ["错误", "false", "no", "否", "0"]):
        return 0
    return 1 if fallback else 0


def product_display_name(category: str) -> str:
    mapping = {
        "bottle": "瓶子",
        "capsule": "胶囊",
        "metal_nut": "金属螺母",
    }
    normalized = normalize_category(category)
    return mapping.get(normalized, safe_str(category))


def format_numeric_range(values, decimals=0) -> str:
    nums = []
    for value in values:
        try:
            if safe_str(value) == "":
                continue
            nums.append(float(value))
        except Exception:
            continue
    if not nums:
        return "-"
    fmt = f"{{:.{decimals}f}}"
    lo, hi = min(nums), max(nums)
    if abs(lo - hi) < 1e-9:
        return fmt.format(lo)
    return f"{fmt.format(lo)}-{fmt.format(hi)}"


def build_purchase_standard_rows(trials: list) -> list:
    grouped = {}
    for trial in trials:
        if trial.get("task_type") != "exp2":
            continue
        category = normalize_category(trial.get("category", ""))
        grouped.setdefault(category, []).append(trial)

    rows = []
    ordered_categories = ["bottle", "capsule", "metal_nut"]
    ordered_categories += sorted(c for c in grouped if c not in ordered_categories)
    for category in ordered_categories:
        subset = grouped.get(category, [])
        if not subset:
            continue
        purchase_trials = [t for t in subset if to_int(t.get("true_code")) == 1]
        rows.append(
            {
                "产品": product_display_name(category),
                "质量底线": f"≥ {format_numeric_range([t.get('quality_gate') for t in subset])}",
                "可采购样本质量分": format_numeric_range([t.get("quality_score") for t in purchase_trials]),
                "可采购样本价格": f"{format_numeric_range([t.get('supplier_price') for t in purchase_trials], 1)} 元/件",
                "综合门槛": f"≥ {format_numeric_range([t.get('purchase_threshold') for t in subset])}",
            }
        )
    return rows


def decision_is_correct(decision_code: int, true_code: int) -> int:
    return int(decision_code == true_code)


def calc_dependence(adopted: int, ai_correct: int) -> str:
    if ai_correct == 1:
        return "proper" if adopted == 1 else "under"
    return "over" if adopted == 1 else "proper"


def resolve_image_path(raw_path: str) -> Path:
    if not raw_path:
        raise FileNotFoundError("题库中未提供图片路径。")

    raw_norm = safe_str(raw_path).replace("\\", "/")
    root = Path(DATASET_ROOT)
    direct = Path(raw_norm)
    candidates = []

    def add_candidate(path_obj: Path):
        if path_obj not in candidates:
            candidates.append(path_obj)

    add_candidate(direct)
    add_candidate(root / raw_norm)

    if "00_raw/" in raw_norm:
        rel = raw_norm.split("00_raw/", 1)[1]
        add_candidate(root / rel)

    parts = [p for p in raw_norm.split("/") if p]
    for key in ["bottle", "capsule", "metal_nut"]:
        if key in parts:
            idx = parts.index(key)
            add_candidate(root / Path(*parts[idx:]))
            break

    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return candidate

    filename = Path(raw_norm).name
    if filename and root.exists():
        matches = list(root.rglob(filename))
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raw_lower = raw_norm.lower()
            for match in matches:
                match_str = str(match).replace("\\", "/").lower()
                if all(seg.lower() in match_str for seg in parts[-3:]):
                    return match
            return matches[0]

    raise FileNotFoundError(
        f"图片未找到：{raw_path}。\n请检查图片路径是否仍为 TO_FILL 占位路径，或当前 00_raw 目录是否与题库一致。"
    )


def ensure_results_dir(participant_id: str) -> Path:
    out_dir = Path(RESULTS_DIR_DEFAULT) / participant_id
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def save_json(path: Path, payload: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def map_strategy_code(text: str) -> str:
    mapping = {
        "主要依靠自己的图像判断": "self",
        "图像判断和AI建议各参考一半": "half",
        "主要参考AI建议": "ai",
        "视情况而定": "depends",
        "主要看质量，再兼顾成本": "quality_first",
        "质量与成本大致各占一半": "balanced",
        "主要看成本，只要质量别太差": "cost_first",
    }
    return mapping.get(text, safe_str(text))


def map_changed_reason_code(text: str) -> str:
    mapping = {
        "AI建议与我不同，选择相信AI": "trust_ai",
        "AI的解释让我重新审视图像": "recheck_image",
        "不确定时偏向跟随AI": "uncertain_follow_ai",
        "我没有改变过判断": "no_change",
        "看到价格后改变了判断": "price_changed",
        "看到质量信息后改变了判断": "quality_changed",
    }
    return mapping.get(text, safe_str(text))


def load_all_banks():
    workbook_path = find_workbook_path()
    try:
        exp1_df = read_structured_sheet(workbook_path, "题库总表")
        exp2_df = read_structured_sheet(workbook_path, "Exp2_多目标采购")
        practice_df = read_structured_sheet(workbook_path, "Practice")
    except Exception as e:
        st.error(f"题库读取失败：{e}")
        return workbook_path, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    return workbook_path, exp1_df, exp2_df, practice_df


def shuffle_rows(df: pd.DataFrame, salt: str) -> pd.DataFrame:
    indices = list(df.index)
    rnd = random.Random(stable_hash_int(salt))
    rnd.shuffle(indices)
    return df.loc[indices].reset_index(drop=True)


def split_quality_bank(df: pd.DataFrame, participant_key: str):
    """拆分实验一/实验二共用质检题库，保证两实验不重复且各自平衡。"""
    work = df.copy()
    work["_category_norm"] = work["产品类别"].apply(normalize_category)
    work["_label_norm"] = work["真实标签"].apply(normalize_okng_label)

    exp1_parts = []
    exp2_parts = []
    for category in PRODUCT_CATEGORIES:
        for label in ("OK", "NG"):
            subset = work[(work["_category_norm"] == category) & (work["_label_norm"] == label)].copy()
            if len(subset) < 8:
                raise ValueError(
                    f"质检题库中 {product_display_name(category)} / {label} 题量不足 8 道，"
                    "无法拆分为实验一和实验二各 4 道。"
                )
            subset = shuffle_rows(subset, f"{participant_key}_quality_split_{category}_{label}")
            exp1_parts.append(subset.iloc[:4])
            exp2_parts.append(subset.iloc[4:8])

    helper_cols = ["_category_norm", "_label_norm"]
    exp1_df = pd.concat(exp1_parts, ignore_index=True).drop(columns=helper_cols)
    exp2_df = pd.concat(exp2_parts, ignore_index=True).drop(columns=helper_cols)
    return exp1_df, exp2_df


def sample_purchase_bank(df: pd.DataFrame, participant_key: str) -> pd.DataFrame:
    """从 48 道采购题中抽 24 道：每类产品 8 道，采购/不采购 12/12，AI 正确率 75%。"""
    work = df.copy()
    work["_category_norm"] = work["产品类别"].apply(normalize_category)
    work["_purchase_code"] = work.apply(
        lambda row: to_int(row.get("purchase_gt_code"), purchase_code(row.get("purchase_gt_label"))),
        axis=1,
    )
    work["_ai_correct"] = work.apply(
        lambda row: to_int(
            row.get("ai_correct_code"),
            int(
                to_int(row.get("ai_suggestion_code"), purchase_code(row.get("ai_suggestion_label")))
                == to_int(row.get("purchase_gt_code"), purchase_code(row.get("purchase_gt_label")))
            ),
        ),
        axis=1,
    )

    parts = []
    for category in PRODUCT_CATEGORIES:
        subset = work[work["_category_norm"] == category].copy()

        selected = []
        # 每个产品抽 8 道：采购 4 / 不采购 4，AI 正确 6 / 错误 2。
        # 优先让采购和不采购各含 1 道 AI 错误；若题库分布不允许，再使用可行组合。
        for purchase_wrong_n in [1, 0, 2]:
            plan = {
                (1, 1): 4 - purchase_wrong_n,
                (1, 0): purchase_wrong_n,
                (0, 1): 2 + purchase_wrong_n,
                (0, 0): 2 - purchase_wrong_n,
            }
            candidate = []
            feasible = True
            for (purchase_code_value, ai_correct_value), need_n in plan.items():
                if need_n <= 0:
                    continue
                pool = subset[
                    (subset["_purchase_code"] == purchase_code_value)
                    & (subset["_ai_correct"] == ai_correct_value)
                ].copy()
                if len(pool) < need_n:
                    feasible = False
                    break
                pool = shuffle_rows(
                    pool,
                    f"{participant_key}_purchase_split_{category}_{purchase_code_value}_{ai_correct_value}",
                )
                candidate.append(pool.iloc[:need_n])
            if feasible:
                selected = candidate
                break

        if not selected:
            raise ValueError(
                f"采购题库中 {product_display_name(category)} 的 AI 正确/错误题量不足，"
                "无法抽取采购 4 道、不采购 4 道且 AI 正确率 75% 的正式题。"
            )
        parts.extend(selected)

    return pd.concat(parts, ignore_index=True).drop(columns=["_category_norm", "_purchase_code", "_ai_correct"])


def build_quality_trials(
    df: pd.DataFrame,
    participant_key: str,
    exp_name: str,
    condition: str,
    explanation_mode: str,
    explanation_column: str,
    order_salt: str,
):
    trials = []
    for _, row in df.iterrows():
        true_label = normalize_okng_label(row.get("真实标签"))
        ai_label = normalize_okng_label(row.get("AI建议"))
        trials.append(
            {
                "task_type": "exp1",
                "exp_name": exp_name,
                "trial_id": safe_str(row.get("题号")),
                "item_id": safe_str(row.get("图片ID")),
                "category": safe_str(row.get("产品类别")),
                "defect_type": safe_str(row.get("缺陷类型")),
                "complexity": safe_str(row.get("复杂度代码") or row.get("复杂度")),
                "true_label": true_label,
                "true_code": okng_code(true_label),
                "ai_label": ai_label,
                "ai_code": okng_code(ai_label),
                "ai_correct": parse_ai_correct(
                    row.get("AI是否正确"),
                    fallback=okng_code(ai_label) == okng_code(true_label),
                ),
                "image_path": safe_str(row.get("图片源路径")),
                "explanation_mode": explanation_mode,
                "explanation_text": safe_str(row.get(explanation_column)) if explanation_column else "",
                "condition": condition,
                "ui_decision_labels": ("OK", "NG"),
                "feedback_text": "",
            }
        )
    rnd = random.Random(stable_hash_int(participant_key + order_salt))
    rnd.shuffle(trials)
    return trials


def build_purchase_trials(
    df: pd.DataFrame,
    participant_key: str,
    exp_name: str,
    condition: str,
    explanation_mode: str,
    explanation_column: str,
    order_salt: str,
):
    trials = []
    for _, row in df.iterrows():
        true_label = normalize_purchase_label(row.get("purchase_gt_label"))
        ai_label = normalize_purchase_label(row.get("ai_suggestion_label"))
        trials.append(
            {
                "task_type": "exp2",
                "exp_name": exp_name,
                "trial_id": safe_str(row.get("题号")),
                "item_id": safe_str(row.get("图片ID")),
                "category": safe_str(row.get("产品类别")),
                "defect_type": safe_str(row.get("缺陷类型")),
                "complexity": safe_str(row.get("复杂度代码") or row.get("复杂度")),
                "image_path": safe_str(row.get("图片源路径")),
                "quality_score": to_float(row.get("quality_score")),
                "supplier_price": to_float(row.get("supplier_price")),
                "cost_score": to_float(row.get("cost_score")),
                "quality_gate": to_float(row.get("quality_gate")),
                "quality_weight": to_float(row.get("quality_weight")),
                "cost_weight": to_float(row.get("cost_weight")),
                "weighted_score": to_float(row.get("weighted_score")),
                "purchase_threshold": to_float(row.get("purchase_threshold")),
                "true_label": true_label,
                "true_code": to_int(row.get("purchase_gt_code"), purchase_code(true_label)),
                "ai_label": ai_label,
                "ai_code": to_int(row.get("ai_suggestion_code"), purchase_code(ai_label)),
                "ai_correct": to_int(row.get("ai_correct_code"), parse_ai_correct(row.get("ai_correct_label"))),
                "decision_zone": safe_str(row.get("decision_zone")),
                "suggestion_type": safe_str(row.get("suggestion_type")),
                "explanation_mode": explanation_mode,
                "explanation_text": safe_str(row.get(explanation_column)) if explanation_column else "",
                "condition": condition,
                "ui_decision_labels": ("采购", "不采购"),
                "feedback_text": "",
            }
        )
    rnd = random.Random(stable_hash_int(participant_key + order_salt))
    rnd.shuffle(trials)
    return trials


def build_practice_trials(
    practice_df: pd.DataFrame,
    exp_name: str,
    source_exp_name: str,
    row_offset: int = 0,
    limit: int = PRACTICE_TRIALS_PER_EXPERIMENT,
    quality_df: pd.DataFrame = None,
    purchase_df: pd.DataFrame = None,
    condition: str = "",
    explanation_mode: str = "none",
    explanation_column: str = "",
):
    filtered = practice_df[practice_df["适用实验"] == source_exp_name].copy().reset_index(drop=True)
    filtered = filtered.iloc[row_offset : row_offset + limit].copy().reset_index(drop=True)

    quality_lookup = {}
    if quality_df is not None and not quality_df.empty:
        for _, quality_row in quality_df.iterrows():
            quality_lookup[safe_str(quality_row.get("题号"))] = quality_row

    purchase_lookup = {}
    if purchase_df is not None and not purchase_df.empty:
        for _, purchase_row in purchase_df.iterrows():
            purchase_lookup[safe_str(purchase_row.get("题号"))] = purchase_row

    trials = []
    for _, row in filtered.iterrows():
        task_type = "exp1" if source_exp_name == "实验一" else "exp2"
        if task_type == "exp1":
            source_row = quality_lookup.get(safe_str(row.get("题号映射")))
            true_label = normalize_okng_label(row.get("标准答案"))
            ai_label = normalize_okng_label(row.get("AI建议"))
            trials.append(
                {
                    "task_type": "exp1",
                    "exp_name": exp_name,
                    "trial_id": safe_str(row.get("练习题号")),
                    "item_id": safe_str(row.get("图片ID")),
                    "category": safe_str(row.get("产品类别")),
                    "defect_type": safe_str(row.get("缺陷类型")),
                    "complexity": safe_str(row.get("复杂度")),
                    "true_label": true_label,
                    "true_code": okng_code(true_label),
                    "ai_label": ai_label,
                    "ai_code": okng_code(ai_label),
                    "ai_correct": int(okng_code(ai_label) == okng_code(true_label)),
                    "image_path": safe_str(row.get("图片源路径")),
                    "explanation_mode": explanation_mode,
                    "explanation_text": safe_str(source_row.get(explanation_column))
                    if source_row is not None and explanation_column
                    else "",
                    "condition": condition,
                    "ui_decision_labels": ("OK", "NG"),
                    "feedback_text": safe_str(row.get("反馈文本")),
                    "practice_source_exp": source_exp_name,
                    "practice_source_trial_id": safe_str(row.get("题号映射")),
                    "is_practice": True,
                }
            )
        else:
            feedback_text = safe_str(row.get("反馈文本"))
            source_row = purchase_lookup.get(safe_str(row.get("题号映射")))

            def purchase_value(purchase_column: str, practice_column: str = None):
                if source_row is not None:
                    value = source_row.get(purchase_column)
                    if safe_str(value) != "":
                        return value
                if practice_column:
                    return row.get(practice_column)
                return ""

            true_label_value = purchase_value("purchase_gt_label", "标准答案")
            ai_label_value = purchase_value("ai_suggestion_label", "AI建议")
            trials.append(
                {
                    "task_type": "exp2",
                    "exp_name": exp_name,
                    "trial_id": safe_str(row.get("练习题号")),
                    "item_id": safe_str(row.get("图片ID")),
                    "category": safe_str(purchase_value("产品类别", "产品类别")),
                    "defect_type": safe_str(purchase_value("缺陷类型", "缺陷类型")),
                    "complexity": safe_str(purchase_value("复杂度代码", "复杂度") or purchase_value("复杂度", "复杂度")),
                    "image_path": safe_str(row.get("图片源路径")),
                    "quality_score": to_float(purchase_value("quality_score")),
                    "supplier_price": to_float(purchase_value("supplier_price")),
                    "cost_score": to_float(purchase_value("cost_score")),
                    "quality_gate": to_float(purchase_value("quality_gate")),
                    "quality_weight": to_float(purchase_value("quality_weight")),
                    "cost_weight": to_float(purchase_value("cost_weight")),
                    "weighted_score": to_float(purchase_value("weighted_score")),
                    "purchase_threshold": to_float(purchase_value("purchase_threshold")),
                    "true_label": normalize_purchase_label(true_label_value),
                    "true_code": to_int(purchase_value("purchase_gt_code", "标准答案代码"), purchase_code(true_label_value)),
                    "ai_label": normalize_purchase_label(ai_label_value),
                    "ai_code": to_int(purchase_value("ai_suggestion_code", "AI建议代码"), purchase_code(ai_label_value)),
                    "ai_correct": to_int(
                        purchase_value("ai_correct_code"),
                        int(purchase_code(ai_label_value) == purchase_code(true_label_value)),
                    ),
                    "decision_zone": safe_str(purchase_value("decision_zone")),
                    "suggestion_type": safe_str(purchase_value("suggestion_type")),
                    "explanation_mode": explanation_mode,
                    "explanation_text": safe_str(purchase_value(explanation_column)) if explanation_column else "",
                    "condition": condition,
                    "ui_decision_labels": ("采购", "不采购"),
                    "feedback_text": feedback_text,
                    "practice_task_type": safe_str(row.get("任务类型")),
                    "practice_source_exp": source_exp_name,
                    "practice_source_trial_id": safe_str(row.get("题号映射")),
                    "is_practice": True,
                }
            )
    return trials


def build_experiment_sequence(exp12_df: pd.DataFrame, exp3_df: pd.DataFrame, practice_df: pd.DataFrame, participant_key: str):
    exp1_df, exp2_quality_df = split_quality_bank(exp12_df, participant_key)
    exp3_sample_df = sample_purchase_bank(exp3_df, participant_key)

    exp1_condition = "with_explanation" if stable_hash_int(participant_key + "_exp1_condition") % 2 == 0 else "no_explanation"
    exp1_explanation_column = "实验一-统一解释内容" if exp1_condition == "with_explanation" else ""
    exp1_explanation_mode = "unified" if exp1_condition == "with_explanation" else "none"

    exp2_condition = "metric" if stable_hash_int(participant_key + "_exp2_condition") % 2 == 0 else "reason"
    exp2_explanation_column = "实验二-指标型解释内容" if exp2_condition == "metric" else "实验二-理由型解释内容"

    exp3_condition = "metric" if stable_hash_int(participant_key + "_exp3_condition") % 2 == 0 else "reason"
    exp3_explanation_column = "实验二-指标型解释内容" if exp3_condition == "metric" else "实验二-理由型解释内容"

    exp1_trials = build_quality_trials(
        exp1_df,
        participant_key,
        exp_name="实验一",
        condition=exp1_condition,
        explanation_mode=exp1_explanation_mode,
        explanation_column=exp1_explanation_column,
        order_salt="_exp1_order",
    )
    exp2_trials = build_quality_trials(
        exp2_quality_df,
        participant_key,
        exp_name="实验二",
        condition=exp2_condition,
        explanation_mode=exp2_condition,
        explanation_column=exp2_explanation_column,
        order_salt="_exp2_order",
    )
    exp3_trials = build_purchase_trials(
        exp3_sample_df,
        participant_key,
        exp_name="实验三",
        condition=exp3_condition,
        explanation_mode=exp3_condition,
        explanation_column=exp3_explanation_column,
        order_salt="_exp3_order",
    )

    exp1_practice = build_practice_trials(
        practice_df,
        exp_name="实验一",
        source_exp_name="实验一",
        row_offset=0,
        quality_df=exp12_df,
        condition=exp1_condition,
        explanation_mode=exp1_explanation_mode,
        explanation_column=exp1_explanation_column,
    )
    exp2_practice = build_practice_trials(
        practice_df,
        exp_name="实验二",
        source_exp_name="实验一",
        row_offset=PRACTICE_TRIALS_PER_EXPERIMENT,
        quality_df=exp12_df,
        condition=exp2_condition,
        explanation_mode=exp2_condition,
        explanation_column=exp2_explanation_column,
    )
    exp3_practice = build_practice_trials(
        practice_df,
        exp_name="实验三",
        source_exp_name="实验二",
        row_offset=0,
        purchase_df=exp3_df,
        condition=exp3_condition,
        explanation_mode=exp3_condition,
        explanation_column=exp3_explanation_column,
    )

    sequence = [
        {
            "meta": {
                "exp_id": "exp1",
                "exp_name": "实验一",
                "task_family": "quality",
                "design": "有无解释",
                "condition": exp1_condition,
                "formal_n": len(exp1_trials),
                "practice_n": len(exp1_practice),
            },
            "trials": exp1_trials,
            "practice_trials": exp1_practice,
        },
        {
            "meta": {
                "exp_id": "exp2",
                "exp_name": "实验二",
                "task_family": "quality",
                "design": "解释形式",
                "condition": exp2_condition,
                "formal_n": len(exp2_trials),
                "practice_n": len(exp2_practice),
            },
            "trials": exp2_trials,
            "practice_trials": exp2_practice,
        },
        {
            "meta": {
                "exp_id": "exp3",
                "exp_name": "实验三",
                "task_family": "purchase",
                "design": "多目标采购",
                "condition": exp3_condition,
                "formal_n": len(exp3_trials),
                "practice_n": len(exp3_practice),
            },
            "trials": exp3_trials,
            "practice_trials": exp3_practice,
        },
    ]

    for item in sequence:
        meta = item["meta"]
        if len(item["trials"]) != FORMAL_TRIALS_PER_EXPERIMENT:
            raise ValueError(f"{meta['exp_name']} 正式题数量为 {len(item['trials'])}，不是 24。")
        if len(item["practice_trials"]) != PRACTICE_TRIALS_PER_EXPERIMENT:
            raise ValueError(f"{meta['exp_name']} 练习题数量为 {len(item['practice_trials'])}，不是 2。")

    exp1_ids = {trial["trial_id"] for trial in exp1_trials}
    exp2_ids = {trial["trial_id"] for trial in exp2_trials}
    if exp1_ids & exp2_ids:
        raise ValueError("实验一与实验二正式题存在重复，请检查抽题逻辑。")

    return sequence


def experiment_sequence_meta() -> list:
    return [item.get("meta", {}) for item in st.session_state.get("experiment_sequence", [])]


def activate_experiment(index: int):
    sequence = st.session_state.get("experiment_sequence", [])
    if not sequence or index < 0 or index >= len(sequence):
        st.error("实验序列未正确初始化，请返回首页重新开始。")
        st.stop()
    current = sequence[index]
    st.session_state["experiment_index"] = index
    st.session_state["exp_meta"] = current["meta"]
    st.session_state["trials"] = current["trials"]
    st.session_state["practice_trials"] = current["practice_trials"]
    st.session_state["current_index"] = 0
    st.session_state["current_render_id"] = None
    st.session_state["trial_start_ts"] = None
    st.session_state["exp_start_ts"] = None
    st.session_state["trial_phase"] = "initial"
    st.session_state["initial_decision_label"] = None
    st.session_state["initial_decision_code"] = None
    st.session_state["initial_rt_ms"] = None
    st.session_state["show_practice_feedback"] = False
    st.session_state["practice_feedback_text"] = ""


def complete_current_experiment():
    sequence = st.session_state.get("experiment_sequence", [])
    current_index = st.session_state.get("experiment_index", 0)
    current_meta = st.session_state.get("exp_meta", {})
    st.session_state["completed_experiment_name"] = current_meta.get("exp_name", f"实验{current_index + 1}")
    st.session_state["current_index"] = 0
    st.session_state["current_render_id"] = None
    st.session_state["trial_phase"] = "initial"
    st.session_state["exp_start_ts"] = None

    if current_index + 1 < len(sequence):
        next_meta = sequence[current_index + 1].get("meta", {})
        st.session_state["next_experiment_name"] = next_meta.get("exp_name", f"实验{current_index + 2}")
        st.session_state["stage"] = "rest"
    else:
        st.session_state["next_experiment_name"] = ""
        st.session_state["stage"] = "questionnaire"


def participant_sheet_df() -> pd.DataFrame:
    meta = st.session_state.get("participant_meta", {})
    exp_meta = st.session_state.get("exp_meta", {})
    workbook_path = st.session_state.get("workbook_path", "")
    if not meta:
        return pd.DataFrame()
    row = {
        "participant_id": meta.get("participant_id", ""),
        "name": meta.get("name", ""),
        "student_id": meta.get("student_id", ""),
        "age": meta.get("age", ""),
        "gender": meta.get("gender", ""),
        "major": meta.get("major", ""),
        "exp_type": meta.get("exp_type", ""),
        "exp_condition": meta.get("exp_condition", ""),
        "experiment_sequence": meta.get("experiment_sequence", ""),
        "current_exp_name": exp_meta.get("exp_name", ""),
        "design": exp_meta.get("design", ""),
        "workbook_path": workbook_path,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    return pd.DataFrame([row])


def questionnaire_sheet_df() -> pd.DataFrame:
    q = st.session_state.get("questionnaire", {})
    meta = st.session_state.get("participant_meta", {})
    if not q or not meta:
        return pd.DataFrame()
    row = {"participant_id": meta.get("participant_id", "")}
    row.update(q)
    return pd.DataFrame([row])


def trial_sheet_df() -> pd.DataFrame:
    rows = st.session_state.get("responses", [])
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


def save_progress():
    meta = st.session_state.get("participant_meta", {})
    if not meta:
        return
    out_dir = ensure_results_dir(meta.get("participant_id", "unknown"))
    xlsx_path = out_dir / OUTPUT_XLSX_NAME

    participant_df = participant_sheet_df()
    questionnaire_df = questionnaire_sheet_df()
    trial_df = trial_sheet_df()

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        participant_df.to_excel(writer, sheet_name="participant_info", index=False)
        (questionnaire_df if not questionnaire_df.empty else pd.DataFrame(columns=["participant_id"])).to_excel(
            writer, sheet_name="questionnaire", index=False
        )
        (trial_df if not trial_df.empty else pd.DataFrame(columns=["participant_id", "trial_id"])).to_excel(
            writer, sheet_name="trial_data", index=False
        )

    save_json(
        out_dir / "session_meta.json",
        {
            "participant_meta": meta,
            "exp_meta": st.session_state.get("exp_meta", {}),
            "experiment_sequence_meta": experiment_sequence_meta(),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "xlsx_path": str(xlsx_path),
        },
    )


def build_result_workbook_bytes() -> bytes:
    participant_df = participant_sheet_df()
    questionnaire_df = questionnaire_sheet_df()
    trial_df = trial_sheet_df()

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        participant_df.to_excel(writer, sheet_name="participant_info", index=False)
        (questionnaire_df if not questionnaire_df.empty else pd.DataFrame(columns=["participant_id"])).to_excel(
            writer, sheet_name="questionnaire", index=False
        )
        (trial_df if not trial_df.empty else pd.DataFrame(columns=["participant_id", "trial_id"])).to_excel(
            writer, sheet_name="trial_data", index=False
        )
    buffer.seek(0)
    return buffer.getvalue()


def init_session():
    defaults = {
        "stage": "setup",
        "workbook_path": "",
        "participant_meta": {},
        "exp_meta": {},
        "experiment_sequence": [],
        "experiment_index": 0,
        "trials": [],
        "practice_trials": [],
        "current_index": 0,
        "current_render_id": None,
        "trial_start_ts": None,
        "exp_start_ts": None,
        "responses": [],
        "questionnaire": {},
        "rest_done": False,
        "trial_phase": "initial",
        "initial_decision_label": None,
        "initial_decision_code": None,
        "initial_rt_ms": None,
        "resolved_image_path": "",
        "show_debug": False,
        "show_practice_feedback": False,
        "practice_feedback_text": "",
        "completed_experiment_name": "",
        "next_experiment_name": "",
        "final_saved": False,
        "google_uploaded": False,
        "google_upload_error": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_experiment():
    for key in list(st.session_state.keys()):
        if key in {
            "final_saved",
            "google_uploaded",
            "google_upload_error",
            "stage", "workbook_path", "participant_meta", "exp_meta", "experiment_sequence", "experiment_index",
            "trials", "practice_trials",
            "current_index", "current_render_id", "trial_start_ts", "exp_start_ts", "responses",
            "questionnaire", "rest_done", "trial_phase", "initial_decision_label", "initial_decision_code",
            "initial_rt_ms", "resolved_image_path", "show_debug", "show_practice_feedback", "practice_feedback_text",
            "completed_experiment_name", "next_experiment_name", "google_uploaded", "google_upload_message"
        }:
            del st.session_state[key]
    init_session()


def render_setup(exp1_df: pd.DataFrame, exp2_df: pd.DataFrame, practice_df: pd.DataFrame):
    st.title(APP_TITLE)
    st.info(
        "**参与须知：** 本次实验包含三个连续实验：实验一“有无解释”、实验二“解释形式”、"
        "实验三“多目标采购”。请按页面顺序完成全部三个实验，并如实填写以下基本信息。"
    )

    workbook_path = st.session_state.get("workbook_path", "") or find_workbook_path()
    st.caption(f"当前题库文件：{workbook_path}")

    with st.form("setup_form"):
        st.markdown("#### 被试基本信息")
        c1, c2, c3 = st.columns(3)
        with c1:
            name = st.text_input("姓名 *")
            student_id = st.text_input("学号 *")
        with c2:
            age = st.text_input("年龄 *")
            gender = st.selectbox("性别 *", ["", "女", "男"])
        with c3:
            major = st.text_input("专业 *")

        submitted = st.form_submit_button("确认并进入实验", use_container_width=True, type="primary")

    if submitted:
        errors = []
        if not safe_str(name):
            errors.append("姓名")
        if not safe_str(student_id):
            errors.append("学号")
        if not safe_str(age):
            errors.append("年龄")
        if not gender:
            errors.append("性别")
        if not safe_str(major):
            errors.append("专业")
        if errors:
            st.error(f"请填写以下必填项：{'、'.join(errors)}")
            return

        if exp1_df.empty or exp2_df.empty or practice_df.empty:
            st.error("题库未成功加载完整，无法生成三个实验。")
            return

        participant_id = build_participant_id(student_id, "all3")
        participant_key = safe_str(student_id) or participant_id

        try:
            sequence = build_experiment_sequence(exp1_df, exp2_df, practice_df, participant_key)
        except Exception as e:
            st.error(f"实验题目抽取失败：{e}")
            return

        sequence_summary = "；".join(
            f"{item['meta']['exp_name']}={item['meta']['condition']}" for item in sequence
        )

        st.session_state["participant_meta"] = {
            "participant_id": participant_id,
            "name": safe_str(name),
            "student_id": safe_str(student_id),
            "age": safe_str(age),
            "gender": gender,
            "major": safe_str(major),
            "exp_type": "three_experiments",
            "exp_condition": sequence_summary,
            "experiment_sequence": "实验一 -> 实验二 -> 实验三",
        }
        st.session_state["experiment_sequence"] = sequence
        activate_experiment(0)
        st.session_state["current_index"] = 0
        st.session_state["responses"] = []
        st.session_state["questionnaire"] = {}
        st.session_state["rest_done"] = False
        st.session_state["exp_start_ts"] = None
        st.session_state["show_practice_feedback"] = False
        st.session_state["practice_feedback_text"] = ""
        save_progress()
        st.session_state["stage"] = "consent"
        st.rerun()

    with st.expander("查看题库摘要（研究者用）", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**质检题库题量**")
            if not exp1_df.empty:
                st.dataframe(
                    exp1_df.groupby(["产品类别", "真实标签"]).size().reset_index(name="数量"),
                    hide_index=True,
                    use_container_width=True,
                )
        with c2:
            st.markdown("**采购题库题量**")
            if not exp2_df.empty:
                st.dataframe(
                    exp2_df.groupby(["产品类别", "purchase_gt_label"]).size().reset_index(name="数量"),
                    hide_index=True,
                    use_container_width=True,
                )
        with c3:
            st.markdown("**当前设计**")
            st.write(f"每个实验正式题：{FORMAL_TRIALS_PER_EXPERIMENT} 道")
            st.write(f"每个实验练习题：{PRACTICE_TRIALS_PER_EXPERIMENT} 道")
            st.write("实验一与实验二正式题互不重复。")


def render_consent():
    st.title("知情同意")
    st.markdown(CONSENT_TEXT)
    agree = st.checkbox("我已仔细阅读以上内容，自愿参加本实验。")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("返回上一页", use_container_width=True):
            st.session_state["stage"] = "setup"
            st.rerun()
    with c2:
        if st.button("同意并继续", type="primary", use_container_width=True, disabled=not agree):
            st.session_state["stage"] = "instruction"
            st.rerun()


def render_instruction():
    exp_meta = st.session_state.get("exp_meta", {})
    exp_name = exp_meta.get("exp_name", "当前实验")
    task_family = exp_meta.get("task_family", "quality")
    exp_no = st.session_state.get("experiment_index", 0) + 1
    exp_total = max(len(st.session_state.get("experiment_sequence", [])), 1)

    st.title(f"{exp_name}说明")
    st.caption(f"实验进度：{exp_no} / {exp_total}")

    if task_family == "quality":
        st.markdown("---")
        st.markdown("### 一、实验任务")
        st.markdown(
            """
你将看到一系列工业产品图像，每张图片展示的是 **瓶子、胶囊或金属螺母** 之一。
你的任务是判断图中产品是否合格：

| 判定 | 含义 |
|------|------|
| ✅ **OK（合格）** | 产品外观无明显缺陷，可以出厂 |
| ❌ **NG（不合格）** | 产品存在可见异常，不可出厂 |
            """
        )

        st.markdown("---")
        st.markdown("### 二、各产品合格 / 不合格标准")
        for _, info in PRODUCT_STANDARDS.items():
            with st.expander(f"📦 {info['name']}", expanded=True):
                st.markdown(f"**✅ 合格品：** {info['ok']}")
                st.markdown("**❌ 不合格品常见缺陷：**")
                for d in info["ng_types"]:
                    st.markdown(f"- {d}")

        st.markdown("---")
        st.markdown("### 三、每道题的作答流程")
        st.markdown(
            """
每道题分为 **两个步骤**：

**第一步 — 独立判断（看不到 AI 建议）**
> 请先仔细观察产品图像，根据自己的判断选择 OK 或 NG。

**第二步 — 参考 AI 后最终决策**
> 完成初步判断后，系统会显示 AI 的检测结果与可能的解释信息。
> 请综合图像与 AI 建议，给出你的最终判断。最终判断可与初步判断相同或不同。
            """
        )
    else:
        st.markdown("---")
        st.markdown("### 一、实验任务")
        st.markdown(
            """
你将看到一系列工业产品图像，并同时获得该产品的**质量信息**与**供应价格信息**。
你的任务不是单纯判断“合格 / 不合格”，而是结合**质量**与**成本**，判断该产品是否值得采购：

| 判定 | 含义 |
|------|------|
| ✅ **采购** | 该产品整体采购价值较高，建议采购 |
| ❌ **不采购** | 该产品整体采购价值不足，不建议采购 |
            """
        )

        st.markdown("---")
        st.markdown("### 二、采购判断标准")
        st.markdown(
            """
- 质量分越高越好，供应价格越低越好。
- 质量分低于基本门槛时，即使价格较低，也不建议采购。
- 质量达标后，再综合考虑成本条件；本题库中质量权重为 70%，成本权重为 30%。
- 综合分达到采购门槛时，标准答案为“采购”；否则为“不采购”。
            """
        )
        standard_rows = build_purchase_standard_rows(st.session_state.get("trials", []))
        if standard_rows:
            st.table(pd.DataFrame(standard_rows))
            st.caption("表中“可采购样本质量分/价格”为当前题库里标准答案为“采购”的参考范围，用于帮助理解判断标准。")

        st.markdown("---")
        st.markdown("### 三、每道题的作答流程")
        st.markdown(
            """
每道题分为 **两个步骤**：

**第一步 — 独立判断（看不到 AI 建议）**
> 先根据图像、质量信息和供应价格，独立判断“采购”或“不采购”。

**第二步 — 参考 AI 后最终决策**
> 完成初步判断后，系统会显示 AI 的采购建议与解释信息。
> 请综合你自己的判断与 AI 建议，给出最终判断。最终判断可与初步判断相同或不同。
            """
        )

    st.markdown("---")
    st.markdown(
        f"""
**⚠️ 注意事项**
- 本阶段为 **{exp_name}**。
- 前 **{len(st.session_state.get("practice_trials", []))} 题** 为练习题，不计入正式数据。
- 正式题共 **{len(st.session_state.get("trials", []))} 题**。
- 系统会记录你的两次判断和反应时间。
- 练习题会显示标准答案反馈，正式题不会显示。
- 完成当前正式实验后，系统会提示接下来进入下一实验，并建议休息 1–2 分钟。
        """
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("返回知情同意", use_container_width=True):
            st.session_state["stage"] = "consent"
            st.rerun()
    with c2:
        if st.button(f"进入{exp_name}练习题", type="primary", use_container_width=True):
            st.session_state["stage"] = "practice"
            st.session_state["current_index"] = 0
            st.session_state["current_render_id"] = None
            st.rerun()


def render_rest():
    completed = st.session_state.get("completed_experiment_name", "当前实验")
    next_name = st.session_state.get("next_experiment_name", "下一实验")
    st.title("请稍作休息")
    st.success(f"你已完成{completed}。")
    st.markdown(f"接下来你将开始 **{next_name}**，请休息 1–2 分钟后再开始。")
    if st.button(f"开始{next_name}", type="primary", use_container_width=True):
        next_index = st.session_state.get("experiment_index", 0) + 1
        activate_experiment(next_index)
        st.session_state["stage"] = "instruction"
        st.rerun()


def render_exp2_info(trial: dict, mode: str):
    st.markdown("### 当前产品信息")
    quality_score = to_float(trial.get("quality_score", 0))
    supplier_price = to_float(trial.get("supplier_price", 0))
    quality_gate = to_float(trial.get("quality_gate", 0))
    purchase_threshold = to_float(trial.get("purchase_threshold", 0))
    if mode == "practice" and quality_score == 0 and supplier_price == 0:
        st.info("本题为练习题，请结合图像与页面信息练习“采购 / 不采购”判断。")
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.metric("质量分", f"{quality_score:.0f}")
        with c2:
            st.metric("供应价格", f"{supplier_price:.1f} 元/件")
        if quality_gate or purchase_threshold:
            st.caption(
                f"参考：质量分低于 {quality_gate:.0f} 通常不采购；质量达标后结合价格判断，综合门槛为 {purchase_threshold:.0f}。"
            )


def render_practice_feedback():
    st.markdown("---")
    st.markdown("### 练习题反馈")
    st.info(st.session_state.get("practice_feedback_text", ""))
    if st.button("继续下一题", type="primary", use_container_width=True):
        st.session_state["show_practice_feedback"] = False
        st.session_state["practice_feedback_text"] = ""
        st.session_state["current_index"] += 1
        st.session_state["current_render_id"] = None
        st.session_state["trial_phase"] = "initial"
        st.rerun()


def render_trial(trials: list, mode: str):
    idx = st.session_state["current_index"]
    total = len(trials)
    exp_meta = st.session_state.get("exp_meta", {})
    exp_name = exp_meta.get("exp_name", "")
    exp_no = st.session_state.get("experiment_index", 0) + 1
    exp_total = max(len(st.session_state.get("experiment_sequence", [])), 1)

    if idx >= total:
        st.session_state["current_index"] = 0
        st.session_state["current_render_id"] = None
        st.session_state["trial_phase"] = "initial"
        if mode == "practice":
            st.session_state["stage"] = "formal"
        else:
            complete_current_experiment()
        st.rerun()

    trial = trials[idx]
    render_uid = f"{exp_meta.get('exp_id', 'exp')}_{mode}_{trial['task_type']}_{trial['trial_id']}_{idx}"

    if st.session_state["current_render_id"] != render_uid:
        st.session_state["current_render_id"] = render_uid
        st.session_state["trial_start_ts"] = time.time()
        st.session_state["trial_phase"] = "initial"
        st.session_state["initial_decision_label"] = None
        st.session_state["initial_decision_code"] = None
        st.session_state["initial_rt_ms"] = None
        st.session_state["resolved_image_path"] = ""
        st.session_state["show_practice_feedback"] = False
        st.session_state["practice_feedback_text"] = ""

    if mode == "formal" and not st.session_state.get("exp_start_ts"):
        st.session_state["exp_start_ts"] = time.time()

    if mode == "formal":
        elapsed = int(time.time() - st.session_state["exp_start_ts"])
        em, es = elapsed // 60, elapsed % 60
        st.markdown(
            f"<div style='text-align:center;padding:6px 0;font-size:0.95rem;'>"
            f"⏱️ 已用时 <b>{em:02d}:{es:02d}</b> &nbsp;|&nbsp; "
            f"{exp_name} <b>{exp_no} / {exp_total}</b> &nbsp;|&nbsp; 正式题 <b>{idx + 1} / {total}</b>"
            f"</div>",
            unsafe_allow_html=True,
        )
        st.progress((idx + 1) / total)
    else:
        st.progress((idx + 1) / total, text=f"{exp_name}练习题进度：{idx + 1}/{total}")

    st.subheader(f"{exp_name}{'练习题' if mode == 'practice' else '正式题'} {idx + 1} / {total}")

    if mode == "practice" and st.session_state.get("show_practice_feedback"):
        render_practice_feedback()
        return

    phase = st.session_state.get("trial_phase", "initial")
    c1, c2 = st.columns([1.25, 1.0])

    with c1:
        try:
            img_path = resolve_image_path(trial["image_path"])
            st.session_state["resolved_image_path"] = str(img_path)
            img_bytes = load_image_bytes(str(img_path))
            st.image(img_bytes, use_container_width=True)
            if st.session_state.get("show_debug"):
                st.caption(f"图片路径：{img_path}")
        except Exception as e:
            st.error(f"图片读取失败：{e}")

    with c2:
        if trial["task_type"] == "exp2":
            render_exp2_info(trial, mode)
            st.markdown("---")

        if phase == "initial":
            st.markdown("### 第一步：请先独立判断")
            if trial["task_type"] == "exp1":
                st.info("请仔细观察图像，在看到 AI 建议之前，先给出你的初步判断。")
            else:
                st.info("请先根据图像、质量信息和价格信息，独立判断是否采购。")

            labels = trial["ui_decision_labels"]
            col_a, col_b = st.columns(2)

            def submit_initial(label: str, code: int):
                rt = int((time.time() - st.session_state["trial_start_ts"]) * 1000)
                st.session_state["initial_decision_label"] = label
                st.session_state["initial_decision_code"] = code
                st.session_state["initial_rt_ms"] = rt
                st.session_state["trial_phase"] = "final"
                st.session_state["trial_start_ts"] = time.time()
                st.rerun()

            with col_a:
                if st.button(f"✅ {labels[0]}", key=f"i_a_{render_uid}", use_container_width=True):
                    submit_initial(labels[0], 1)
            with col_b:
                if st.button(f"❌ {labels[1]}", key=f"i_b_{render_uid}", use_container_width=True):
                    submit_initial(labels[1], 0)

            st.caption("完成初步判断后，将显示 AI 建议，再进行最终决策。")
        else:
            st.markdown("### 第二步：参考 AI 建议，做最终判断")
            st.info(f"**AI 建议：{trial['ai_label']}**")
            if safe_str(trial.get("explanation_text", "")):
                st.write(trial["explanation_text"])
            st.caption(f"你的初步判断：**{st.session_state['initial_decision_label']}**")
            st.markdown("---")
            st.markdown("**你的最终判断：**")

            labels = trial["ui_decision_labels"]
            col_a, col_b = st.columns(2)

            def submit_final(label: str, code: int):
                rt = int((time.time() - st.session_state["trial_start_ts"]) * 1000)
                init_label = st.session_state["initial_decision_label"]
                init_code = st.session_state["initial_decision_code"]
                init_rt = st.session_state["initial_rt_ms"] or 0
                meta = st.session_state["participant_meta"]
                exp_meta = st.session_state["exp_meta"]

                adopted = int(code == trial["ai_code"])

                record = {
                    "participant_id": meta.get("participant_id", ""),
                    "exp_type": meta.get("exp_type", ""),
                    "participant_condition_summary": meta.get("exp_condition", ""),
                    "exp_id": exp_meta.get("exp_id", ""),
                    "exp_name": exp_meta.get("exp_name", trial.get("exp_name", "")),
                    "exp_sequence_index": st.session_state.get("experiment_index", 0) + 1,
                    "exp_condition": exp_meta.get("condition", trial.get("condition", "")),
                    "design": exp_meta.get("design", ""),
                    "task_family": exp_meta.get("task_family", ""),
                    "task_type": trial["task_type"],
                    "trial_stage": mode,
                    "trial_index": idx + 1,
                    "global_formal_index": len(st.session_state.get("responses", [])) + 1 if mode != "practice" else "",
                    "trial_id": trial["trial_id"],
                    "item_id": trial["item_id"],
                    "category": normalize_category(trial["category"]),
                    "defect_type": safe_str(trial["defect_type"]),
                    "complexity": normalize_complexity(trial["complexity"]),
                    "true_label": trial["true_label"],
                    "true_code": trial["true_code"],
                    "ai_label": trial["ai_label"],
                    "ai_code": trial["ai_code"],
                    "ai_correct": trial["ai_correct"],
                    "explanation_mode": trial["explanation_mode"],
                    "explanation_text": safe_str(trial.get("explanation_text", "")),
                    "initial_decision_label": init_label,
                    "initial_decision_code": init_code,
                    "final_decision_label": label,
                    "final_decision_code": code,
                    "initial_correct": decision_is_correct(init_code, trial["true_code"]),
                    "final_correct": decision_is_correct(code, trial["true_code"]),
                    "initial_rt_ms": init_rt,
                    "final_rt_ms": rt,
                    "total_rt_ms": init_rt + rt,
                    "decision_changed": int(code != init_code),
                    "adopt_ai": adopted,
                    "dependence_type": calc_dependence(adopted, trial["ai_correct"]),
                    "recorded_at": datetime.now().isoformat(timespec="seconds"),
                }

                if trial["task_type"] == "exp2":
                    record.update(
                        {
                            "quality_score": trial.get("quality_score", ""),
                            "supplier_price": trial.get("supplier_price", ""),
                            "cost_score": trial.get("cost_score", ""),
                            "quality_gate": trial.get("quality_gate", ""),
                            "quality_weight": trial.get("quality_weight", ""),
                            "cost_weight": trial.get("cost_weight", ""),
                            "weighted_score": trial.get("weighted_score", ""),
                            "purchase_threshold": trial.get("purchase_threshold", ""),
                            "decision_zone": trial.get("decision_zone", ""),
                            "suggestion_type": trial.get("suggestion_type", ""),
                        }
                    )

                if mode != "practice":
                    st.session_state["responses"].append(record)
                    # 只在每个实验结束时保存一次，避免每题写 Excel 导致卡顿
                    if idx + 1 >= total:
                        save_progress()
                    st.session_state["current_index"] += 1
                    st.session_state["current_render_id"] = None
                    st.session_state["trial_phase"] = "initial"
                    st.rerun()
                else:
                    st.session_state["practice_feedback_text"] = trial.get("feedback_text", "")
                    st.session_state["show_practice_feedback"] = True
                    st.rerun()

            with col_a:
                if st.button(f"✅ {labels[0]}", key=f"f_a_{render_uid}", use_container_width=True):
                    submit_final(labels[0], 1)
            with col_b:
                if st.button(f"❌ {labels[1]}", key=f"f_b_{render_uid}", use_container_width=True):
                    submit_final(labels[1], 0)

            st.caption("点击后自动进入下一题。" if mode != "practice" else "练习题会显示标准答案反馈。")


def render_questionnaire():
    st.title("实验结束问卷")
    st.markdown("你已完成三个实验。请根据整个实验过程中的真实感受作答，没有对错之分。")

    with st.form("questionnaire_form"):
        st.markdown("### 第一部分：AI 信任量表")
        st.caption("1=非常不同意，7=非常同意")
        stias_confidence = st.slider("1. 我对该 AI 系统有信心。", 1, 7, 4)
        stias_reliable = st.slider("2. 该 AI 系统是可靠的。", 1, 7, 4)
        stias_trust = st.slider("3. 我可以信任该 AI 系统。", 1, 7, 4)

        st.markdown("### 第二部分：解释满意度量表")
        st.caption("请根据你在实验二和实验三中看到 AI 解释时的整体感受作答。1=非常不同意，5=非常同意")
        ess_understand = st.slider("4. AI 的解释让我能够理解其判断依据。", 1, 5, 3)
        ess_satisfy = st.slider("5. 总体来说，我对 AI 提供的解释感到满意。", 1, 5, 3)
        ess_detail = st.slider("6. AI 的解释提供了足够的细节。", 1, 5, 3)
        ess_complete = st.slider("7. AI 的解释覆盖了我作出判断所需的主要信息。", 1, 5, 3)
        ess_useful = st.slider("8. AI 的解释对我完成判断任务是有帮助的。", 1, 5, 3)
        ess_accuracy = st.slider("9. AI 的解释看起来与图像或产品信息相符合。", 1, 5, 3)
        ess_trust = st.slider("10. AI 的解释让我更愿意相信 AI 的建议。", 1, 5, 3)

        st.markdown("### 第三部分：任务负荷量表（NASA-TLX）")
        st.caption("0=非常低，100=非常高。表现维度中，0=非常成功，100=非常失败。")
        nasa_mental = st.slider("11. 脑力需求：完成任务需要多少脑力投入？", 0, 100, 50)
        nasa_physical = st.slider("12. 体力需求：完成任务需要多少身体操作负担？", 0, 100, 10)
        nasa_temporal = st.slider("13. 时间压力：你感受到多大的时间压力？", 0, 100, 50)
        nasa_performance = st.slider("14. 任务表现：你认为自己完成任务的成功程度如何？", 0, 100, 50)
        nasa_effort = st.slider("15. 努力程度：你需要付出多少努力来完成任务？", 0, 100, 50)
        nasa_frustration = st.slider("16. 挫败感：你在任务中感到多少挫败、烦躁或压力？", 0, 100, 50)

        st.markdown("### 第四部分：任务策略补充")
        quality_strategy = st.radio(
            "17. 在实验一和实验二的质检判断中，你通常的策略是？",
            ["主要依靠自己的图像判断", "图像判断和AI建议各参考一半", "主要参考AI建议", "视情况而定"],
            index=3,
        )
        quality_changed_reason = st.radio(
            "18. 在质检判断中，当你改变了初步判断时，主要原因是？",
            ["AI建议与我不同，选择相信AI", "AI的解释让我重新审视图像", "不确定时偏向跟随AI", "我没有改变过判断"],
            index=3,
        )
        purchase_strategy = st.radio(
            "19. 在实验三的采购判断中，你通常更偏向哪种策略？",
            ["主要看质量，再兼顾成本", "质量与成本大致各占一半", "主要看成本，只要质量别太差", "视情况而定"],
            index=1,
        )
        purchase_changed_reason = st.radio(
            "20. 在采购判断中，当你改变了初步判断时，主要原因是？",
            ["AI建议与我不同，选择相信AI", "看到质量信息后改变了判断", "看到价格后改变了判断", "我没有改变过判断"],
            index=3,
        )
        quality_priority = st.slider("21. 在采购判断中，你认为质量信息的重要性有多高？", 1, 7, 5)
        cost_priority = st.slider("22. 在采购判断中，你认为价格信息的重要性有多高？", 1, 7, 4)
        rule_awareness = st.slider("23. 你是否感觉系统内部存在某种固定的采购判断规则？", 1, 7, 5)

        comments = st.text_area("24. 如有其他想说的（例如：哪些题目较难、对实验的建议等），请在此填写：", placeholder="选填")
        submitted = st.form_submit_button("提交问卷", type="primary", use_container_width=True)

    if submitted:
        stias_items = [stias_confidence, stias_reliable, stias_trust]
        ess_items = [ess_understand, ess_satisfy, ess_detail, ess_complete, ess_useful, ess_accuracy, ess_trust]
        nasa_items = [nasa_mental, nasa_physical, nasa_temporal, nasa_performance, nasa_effort, nasa_frustration]
        st.session_state["questionnaire"] = {
            "scale_version": "S_TIAS_ESS_NASA_TLX_2026",
            "stias_confidence": stias_confidence,
            "stias_reliable": stias_reliable,
            "stias_trust": stias_trust,
            "stias_mean": round(sum(stias_items) / len(stias_items), 3),
            "ess_understand": ess_understand,
            "ess_satisfaction": ess_satisfy,
            "ess_detail": ess_detail,
            "ess_complete": ess_complete,
            "ess_useful": ess_useful,
            "ess_accuracy": ess_accuracy,
            "ess_trust": ess_trust,
            "ess_mean": round(sum(ess_items) / len(ess_items), 3),
            "nasa_mental": nasa_mental,
            "nasa_physical": nasa_physical,
            "nasa_temporal": nasa_temporal,
            "nasa_performance": nasa_performance,
            "nasa_effort": nasa_effort,
            "nasa_frustration": nasa_frustration,
            "nasa_raw_mean": round(sum(nasa_items) / len(nasa_items), 3),
            "quality_strategy": map_strategy_code(quality_strategy),
            "quality_changed_reason": map_changed_reason_code(quality_changed_reason),
            "purchase_strategy": map_strategy_code(purchase_strategy),
            "purchase_changed_reason": map_changed_reason_code(purchase_changed_reason),
            "quality_priority": quality_priority,
            "cost_priority": cost_priority,
            "rule_awareness": rule_awareness,
            "comments": safe_str(comments),
        }
        save_progress()
        st.session_state["stage"] = "finish"
        st.rerun()



def dataframe_to_sheet_values(df: pd.DataFrame) -> list:
    """将 DataFrame 转为 Google Sheets 可写入的二维列表。"""
    if df is None or df.empty:
        return []
    safe_df = df.copy()
    safe_df = safe_df.astype(object).where(pd.notna(safe_df), "")
    return safe_df.astype(str).values.tolist()


def get_google_sheet_config():
    """
    从 Streamlit Secrets 读取 Google Sheets 配置。

    需要在 Streamlit Cloud 的 App secrets 中配置：
    google_sheet_id = "你的表格ID"

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "..."
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."
    """
    try:
        sheet_id = st.secrets.get("google_sheet_id", "")
        service_account_info = st.secrets.get("gcp_service_account", None)
    except Exception:
        return "", None
    return sheet_id, service_account_info


def google_sheets_is_configured() -> bool:
    sheet_id, service_account_info = get_google_sheet_config()
    return bool(gspread is not None and sheet_id and service_account_info)


def get_or_create_worksheet(spreadsheet, sheet_name: str, headers: list):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=sheet_name,
            rows=max(1000, len(headers) + 10),
            cols=max(26, len(headers)),
        )
        worksheet.update("A1", [headers])
        return worksheet

    existing_headers = worksheet.row_values(1)
    if not existing_headers:
        worksheet.update("A1", [headers])
    else:
        missing_headers = [h for h in headers if h not in existing_headers]
        if missing_headers:
            updated_headers = existing_headers + missing_headers
            worksheet.update("A1", [updated_headers])
    return worksheet


def append_df_to_google_sheet(spreadsheet, sheet_name: str, df: pd.DataFrame):
    if df is None or df.empty:
        return
    headers = [str(c) for c in df.columns]
    worksheet = get_or_create_worksheet(spreadsheet, sheet_name, headers)
    sheet_headers = worksheet.row_values(1) or headers
    safe_df = df.copy()
    for header in sheet_headers:
        if header not in safe_df.columns:
            safe_df[header] = ""
    safe_df = safe_df[sheet_headers]
    values = dataframe_to_sheet_values(safe_df)
    if values:
        worksheet.append_rows(values, value_input_option="USER_ENTERED")


def upload_current_result_to_google_sheets() -> tuple[bool, str]:
    """实验完成后，将 participant/questionnaire/trial 三张表追加到同一个 Google Sheet。"""
    if st.session_state.get("google_uploaded"):
        return True, st.session_state.get("google_upload_message", "数据已上传至 Google Sheet。")

    if not google_sheets_is_configured():
        if gspread is None:
            return False, "未安装 gspread，无法自动上传。请在 requirements.txt 中加入 gspread。"
        return False, "尚未配置 Google Sheets 密钥，当前仍可通过页面按钮下载 Excel。"

    try:
        sheet_id, service_account_info = get_google_sheet_config()
        client = gspread.service_account_from_dict(dict(service_account_info))
        spreadsheet = client.open_by_key(sheet_id)

        append_df_to_google_sheet(spreadsheet, "participant_info", participant_sheet_df())
        append_df_to_google_sheet(spreadsheet, "questionnaire", questionnaire_sheet_df())
        append_df_to_google_sheet(spreadsheet, "trial_data", trial_sheet_df())

        msg = "数据已自动上传至研究者的 Google Sheet 总表。"
        st.session_state["google_uploaded"] = True
        st.session_state["google_upload_message"] = msg
        return True, msg
    except Exception as e:
        return False, f"Google Sheet 自动上传失败：{e}。请使用下方按钮下载 Excel 后发给研究者。"


def render_finish():
    st.title("🎉 实验完成")
    st.success("感谢你的参与，三个实验已经全部完成。")
    participant_id = st.session_state["participant_meta"].get("participant_id", "unknown")
    out_dir = ensure_results_dir(participant_id)
    # 关键：避免页面刷新导致重复保存/重复上传
    if "final_saved" not in st.session_state:
        st.session_state["final_saved"] = False
    if "google_uploaded" not in st.session_state:
        st.session_state["google_uploaded"] = False
    if "google_upload_error" not in st.session_state:
        st.session_state["google_upload_error"] = ""
    # 只在第一次进入完成页时保存本地 Excel
    if not st.session_state["final_saved"]:
        save_progress()
        st.session_state["final_saved"] = True
    # 只在第一次进入完成页时上传 Google Sheet
    if not st.session_state["google_uploaded"]:
        try:
            upload_to_google_sheet()
            st.session_state["google_uploaded"] = True
            st.session_state["google_upload_error"] = ""
        except Exception as e:
            st.session_state["google_upload_error"] = str(e)
    if st.session_state["google_uploaded"]:
        st.success("数据已成功上传至 Google Sheet。")
    else:
        st.warning("Google Sheet 自动上传失败。请点击下方按钮下载 Excel，并发送给研究者。")
        st.caption(st.session_state["google_upload_error"])
    st.markdown("### 数据备份")
    st.info("为避免网络波动导致数据丢失，请下载本次实验数据作为备份。")
    st.download_button(
        "⬇️ 下载本次实验数据（Excel）",
        data=build_result_workbook_bytes(),
        file_name=f"{participant_id}_{OUTPUT_XLSX_NAME}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if st.button("🏠 下一位被试（返回首页）", use_container_width=True, type="primary"):
        reset_experiment()
        st.session_state["stage"] = "setup"
        st.rerun()


def render_sidebar():
    with st.sidebar:
        st.title("管理员设置")
        st.session_state["show_debug"] = st.checkbox("显示调试信息", value=False)
        st.write("默认题库会优先读取多目标采购版工作簿。")
        if st.session_state.get("exp_meta"):
            with st.expander("当前会话信息"):
                st.write(st.session_state.get("exp_meta", {}))
                st.write(st.session_state.get("participant_meta", {}))
        st.markdown("---")
        if st.button("重置当前会话", use_container_width=True):
            reset_experiment()
            st.rerun()


def validate_ready(exp1_df: pd.DataFrame, exp2_df: pd.DataFrame, practice_df: pd.DataFrame) -> bool:
    if exp1_df.empty or exp2_df.empty or practice_df.empty:
        st.warning("题库未成功加载完整，请先检查工作簿路径和 sheet 结构。")
        return False
    if not DATASET_ROOT.exists():
        st.warning(f"当前图片根目录不存在：{DATASET_ROOT}。如果图片不显示，请优先检查该路径。")
    return True


def main():
    init_session()
    render_sidebar()
    workbook_path, exp1_df, exp2_df, practice_df = load_all_banks()
    st.session_state["workbook_path"] = workbook_path
    stage = st.session_state["stage"]

    if stage == "setup":
        render_setup(exp1_df, exp2_df, practice_df)
        return

    if not validate_ready(exp1_df, exp2_df, practice_df):
        st.stop()

    if stage == "consent":
        render_consent()
    elif stage == "instruction":
        render_instruction()
    elif stage == "practice":
        render_trial(st.session_state["practice_trials"], "practice")
    elif stage == "formal":
        render_trial(st.session_state["trials"], "formal")
    elif stage == "rest":
        render_rest()
    elif stage == "questionnaire":
        render_questionnaire()
    elif stage == "finish":
        render_finish()


if __name__ == "__main__":
    main()
