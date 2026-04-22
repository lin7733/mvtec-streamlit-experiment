import hashlib
import json
import random
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from PIL import Image, ImageFile
import streamlit as st

ImageFile.LOAD_TRUNCATED_IMAGES = True

st.set_page_config(
    page_title="MVTec AD 解释透明度实验",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_TITLE = "MVTec AD 解释透明度实验平台"
WORKBOOK_PATH = "05_metadata/MVTec_实验题库_完整版_解释优化版.xlsx"
DATASET_ROOT = "00_raw"
RESULTS_DIR_DEFAULT = "results"
PRACTICE_TRIALS = 4
BREAK_AFTER = 24
OUTPUT_XLSX_NAME = "experiment_data.xlsx"

CONSENT_TEXT = """
**知情同意书**

本实验为本科毕业论文研究项目，研究主题为“AI 解释透明度对人机协作决策的影响”。

**实验内容：** 你将完成一系列工业产品质检判断任务，并在实验结束后填写问卷。

**数据使用：** 实验过程中记录的作答结果、反应时间和问卷评分，仅用于学术研究分析。

**自愿参与：** 你可以在任何时候选择退出实验，不会产生任何不良后果。

**实验时长：** 约 20–30 分钟。

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

REQUIRED_COLS = [
    "题号", "图片ID", "产品类别", "缺陷类型", "复杂度", "真实标签", "AI建议", "AI是否正确",
    "实验一-无解释呈现", "实验一-统一解释内容", "实验二-指标型解释内容", "实验二-理由型解释内容", "图片源路径"
]


def stable_hash_int(text: str) -> int:
    return int(hashlib.md5(text.encode("utf-8")).hexdigest(), 16)


def safe_str(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [safe_str(c) for c in df.columns]
    for c in df.columns:
        df[c] = df[c].apply(safe_str)
    return df


def read_bank() -> pd.DataFrame:
    try:
        return normalize_df(pd.read_excel(WORKBOOK_PATH, sheet_name="题库总表"))
    except Exception as e:
        st.error(f"题库读取失败：{e}")
        return pd.DataFrame()


def sanitize_for_path(text: str) -> str:
    cleaned = "".join(ch for ch in safe_str(text) if ch.isalnum() or ch in {"_", "-"})
    return cleaned or "unknown"


def build_participant_id(student_id: str, exp_short: str) -> str:
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    return f"{exp_short}_{sanitize_for_path(student_id)}_{ts}"


def normalize_label(value: str) -> str:
    text = safe_str(value).lower().replace("（", "(").replace("）", ")")
    if any(k in text for k in ["ok", "合格", "正常", "good", "no defect", "无缺陷"]):
        return "OK"
    if any(k in text for k in ["ng", "不合格", "异常", "缺陷", "bad", "defect", "有缺陷"]):
        return "NG"
    return safe_str(value)


def normalize_complexity(value: str) -> str:
    text = safe_str(value).lower()
    if any(k in text for k in ["low", "低"]):
        return "low"
    if any(k in text for k in ["high", "高"]):
        return "high"
    return safe_str(value)


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


def parse_ai_correct(value: str, ai_suggestion: str = "", true_label: str = "") -> int:
    text = safe_str(value).lower()
    if any(k in text for k in ["正确", "true", "yes", "是", "1"]):
        return 1
    if any(k in text for k in ["错误", "false", "no", "否", "0"]):
        return 0
    return int(normalize_label(ai_suggestion) == normalize_label(true_label))


def calc_adoption(final_decision: str, ai_suggestion: str) -> int:
    return int(normalize_label(final_decision) == normalize_label(ai_suggestion))


def calc_dependence_code(final_decision: str, ai_suggestion: str, ai_correct: str, true_label: str) -> str:
    adopted = normalize_label(final_decision) == normalize_label(ai_suggestion)
    ai_is_correct = parse_ai_correct(ai_correct, ai_suggestion, true_label) == 1
    if ai_is_correct:
        return "proper" if adopted else "under"
    return "over" if adopted else "proper"


def decision_is_correct(decision: str, true_label: str) -> int:
    return int(normalize_label(decision) == normalize_label(true_label))


def resolve_image_path(raw_path: str) -> Path:
    if not raw_path:
        raise FileNotFoundError("题库中未提供图片路径。")

    raw_norm = safe_str(raw_path).replace("\\", "/")
    root = Path(DATASET_ROOT)
    direct = Path(raw_norm)
    candidates = []

    def add_candidate(p: Path):
        if p not in candidates:
            candidates.append(p)

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

    for p in candidates:
        if p.exists() and p.is_file():
            return p

    filename = Path(raw_norm).name
    if filename and root.exists():
        matches = list(root.rglob(filename))
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raw_lower = raw_norm.lower()
            for m in matches:
                m_str = str(m).replace("\\", "/").lower()
                if all(seg.lower() in m_str for seg in parts[-3:]):
                    return m
            return matches[0]

    raise FileNotFoundError(
        f"图片未找到：{raw_path}。\n"
        f"请检查图片源路径是否为旧设备绝对路径，或当前 00_raw 目录是否与题库对应。"
    )


def ensure_results_dir(participant_id: str) -> Path:
    out_dir = Path(RESULTS_DIR_DEFAULT) / participant_id
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def save_json(path: Path, payload: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def get_condition_from_participant(participant_key: str, options: list) -> str:
    return options[stable_hash_int(participant_key) % len(options)]


def map_strategy_code(text: str) -> str:
    mapping = {
        "主要依靠自己的图像判断": "self",
        "图像判断和AI建议各参考一半": "half",
        "主要参考AI建议": "ai",
        "视情况而定": "depends",
    }
    return mapping.get(text, safe_str(text))


def map_changed_reason_code(text: str) -> str:
    mapping = {
        "AI建议与我不同，选择相信AI": "trust_ai",
        "AI的解释让我重新审视图像": "recheck_image",
        "不确定时偏向跟随AI": "uncertain_follow_ai",
        "我没有改变过判断": "no_change",
    }
    return mapping.get(text, safe_str(text))


def map_exp2_choice_code(text: str) -> str:
    mapping = {
        "指标型解释（数字/百分比）": "metric",
        "理由型解释（文字描述）": "reason",
        "两者差不多": "same",
    }
    return mapping.get(text, safe_str(text))


def build_exp1_trials(df: pd.DataFrame, participant_key: str):
    condition = get_condition_from_participant(participant_key, ["无解释", "有解释"])
    trials = []
    for _, row in df.iterrows():
        explanation_text = "" if condition == "无解释" else row["实验一-统一解释内容"]
        trials.append({
            "trial_id": row["题号"],
            "item_id": row["图片ID"],
            "category": row["产品类别"],
            "defect_type": row["缺陷类型"],
            "complexity": row["复杂度"],
            "true_label": row["真实标签"],
            "ai_suggestion": row["AI建议"],
            "ai_correct": row["AI是否正确"],
            "image_path": row["图片源路径"],
            "explanation_mode": condition,
            "explanation_text": explanation_text,
            "exp_name": "实验一",
        })
    rnd = random.Random(stable_hash_int(participant_key + "_exp1_order"))
    rnd.shuffle(trials)
    return trials, {"exp_name": "实验一", "design": "between", "condition": condition}


def build_exp2_trials(df: pd.DataFrame, participant_key: str):
    version = "A" if stable_hash_int(participant_key) % 2 == 0 else "B"
    tmp = df.copy()
    tmp["group_key"] = tmp["产品类别"] + "|" + tmp["复杂度"] + "|" + tmp["真实标签"]
    trials = []
    for _, g in tmp.groupby("group_key", sort=True):
        g = g.sort_values("题号").reset_index(drop=True)
        split = len(g) // 2
        metric_idx = set(g.index[:split]) if version == "A" else set(g.index[split:])
        for idx, row in g.iterrows():
            is_metric = idx in metric_idx
            trials.append({
                "trial_id": row["题号"],
                "item_id": row["图片ID"],
                "category": row["产品类别"],
                "defect_type": row["缺陷类型"],
                "complexity": row["复杂度"],
                "true_label": row["真实标签"],
                "ai_suggestion": row["AI建议"],
                "ai_correct": row["AI是否正确"],
                "image_path": row["图片源路径"],
                "explanation_mode": "指标型解释" if is_metric else "理由型解释",
                "explanation_text": row["实验二-指标型解释内容"] if is_metric else row["实验二-理由型解释内容"],
                "exp_name": "实验二",
            })
    rnd = random.Random(stable_hash_int(participant_key + "_exp2_order"))
    rnd.shuffle(trials)
    return trials, {"exp_name": "实验二", "design": "within", "counterbalance_version": version}


def select_practice_trials(trials: list, participant_key: str, n: int):
    if n <= 0:
        return []
    copied = [t.copy() for t in trials]
    rnd = random.Random(stable_hash_int(participant_key + "_practice"))
    rnd.shuffle(copied)
    practice = copied[:min(n, len(copied))]
    for t in practice:
        t["is_practice"] = True
    return practice


def participant_sheet_df() -> pd.DataFrame:
    meta = st.session_state.get("participant_meta", {})
    exp_meta = st.session_state.get("exp_meta", {})
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
        "design": exp_meta.get("design", ""),
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
        if questionnaire_df.empty:
            pd.DataFrame(columns=["participant_id"]).to_excel(writer, sheet_name="questionnaire", index=False)
        else:
            questionnaire_df.to_excel(writer, sheet_name="questionnaire", index=False)
        if trial_df.empty:
            pd.DataFrame(columns=[
                "participant_id", "trial_index", "trial_id", "category", "complexity", "true_label",
                "ai_label", "initial_decision", "final_decision", "total_rt_ms"
            ]).to_excel(writer, sheet_name="trial_data", index=False)
        else:
            trial_df.to_excel(writer, sheet_name="trial_data", index=False)

    save_json(out_dir / "session_meta.json", {
        "participant_meta": meta,
        "exp_meta": st.session_state.get("exp_meta", {}),
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "xlsx_path": str(xlsx_path),
    })


def init_session():
    defaults = {
        "stage": "setup",
        "participant_meta": {},
        "exp_meta": {},
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
        "initial_decision": None,
        "initial_rt_ms": None,
        "resolved_image_path": "",
        "show_debug": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_experiment():
    keys = [
        "stage", "participant_meta", "exp_meta", "trials", "practice_trials", "current_index",
        "current_render_id", "trial_start_ts", "exp_start_ts", "responses", "questionnaire",
        "rest_done", "trial_phase", "initial_decision", "initial_rt_ms", "resolved_image_path", "show_debug"
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]
    init_session()


def render_setup(df: pd.DataFrame):
    st.title(APP_TITLE)
    st.info(
        "**参与须知：** 本实验分为实验一和实验二两种类型，**每位被试只需完成其中一种**。"
        "请根据研究者安排选择实验类型，并如实填写以下基本信息。"
    )

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
            exp_name = st.selectbox("实验类型 *", ["实验一：有无解释", "实验二：解释形式"])

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
        if df.empty:
            st.error("题库加载失败，请联系研究者。")
            return

        exp_short = "exp1" if exp_name.startswith("实验一") else "exp2"
        participant_id = build_participant_id(student_id, exp_short)
        participant_key = safe_str(student_id) or participant_id

        if exp_name.startswith("实验一"):
            trials, meta = build_exp1_trials(df, participant_key)
            exp_condition = "with_expl" if meta.get("condition") == "有解释" else "no_expl"
        else:
            trials, meta = build_exp2_trials(df, participant_key)
            exp_condition = safe_str(meta.get("counterbalance_version", ""))

        st.session_state["participant_meta"] = {
            "participant_id": participant_id,
            "name": safe_str(name),
            "student_id": safe_str(student_id),
            "age": safe_str(age),
            "gender": gender,
            "major": safe_str(major),
            "exp_type": exp_short,
            "exp_condition": exp_condition,
        }
        st.session_state["exp_meta"] = meta
        st.session_state["trials"] = trials
        st.session_state["practice_trials"] = select_practice_trials(trials, participant_key, PRACTICE_TRIALS)
        st.session_state["current_index"] = 0
        st.session_state["responses"] = []
        st.session_state["questionnaire"] = {}
        st.session_state["rest_done"] = False
        st.session_state["exp_start_ts"] = None
        save_progress()
        st.session_state["stage"] = "consent"
        st.rerun()

    with st.expander("查看题库摘要（研究者用）", expanded=False):
        if not df.empty:
            st.dataframe(
                df.groupby(["产品类别", "复杂度", "真实标签"]).size().reset_index(name="数量"),
                hide_index=True,
                use_container_width=True,
            )


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
    st.title("实验说明")

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
    st.markdown("### 二、各产品合格/不合格标准")
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
> 完成初步判断后，系统会显示 AI 的检测结果（部分题目还会附上解释信息）。
> 请综合图像与 AI 建议，给出你的最终判断。最终判断可与初步判断相同或不同。
        """
    )

    st.markdown("---")
    st.markdown(
        f"""
**⚠️ 注意事项**
- 前 **{PRACTICE_TRIALS} 题** 为练习题，不计入正式数据。
- 正式实验共 **48 题**，中途会有一次短暂休息（第 {BREAK_AFTER} 题后）。
- 系统会记录你的两次判断和反应时间。
- 实验结束后请完成问卷。
        """
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("返回知情同意", use_container_width=True):
            st.session_state["stage"] = "consent"
            st.rerun()
    with c2:
        label = "进入练习题" if PRACTICE_TRIALS > 0 else "开始正式实验"
        if st.button(label, type="primary", use_container_width=True):
            st.session_state["stage"] = "practice" if PRACTICE_TRIALS > 0 else "formal"
            st.session_state["current_index"] = 0
            st.session_state["current_render_id"] = None
            st.rerun()


def render_rest():
    st.title("请稍作休息 ☕")
    st.markdown("你已完成前 24 题，建议休息 1–2 分钟后再继续。")
    elapsed = int(time.time() - (st.session_state.get("exp_start_ts") or time.time()))
    st.info(f"当前已用时：{elapsed // 60} 分 {elapsed % 60} 秒")
    if st.button("我已休息好，继续实验", type="primary", use_container_width=True):
        st.session_state["rest_done"] = True
        st.session_state["stage"] = "formal"
        st.session_state["current_render_id"] = None
        st.rerun()


def render_trial(trials: list, mode: str):
    idx = st.session_state["current_index"]
    total = len(trials)

    if idx >= total:
        st.session_state["current_index"] = 0
        st.session_state["stage"] = "formal" if mode == "practice" else "questionnaire"
        st.session_state["current_render_id"] = None
        st.session_state["trial_phase"] = "initial"
        st.rerun()

    if mode == "formal" and idx == BREAK_AFTER and not st.session_state.get("rest_done", False):
        st.session_state["stage"] = "rest"
        st.rerun()

    trial = trials[idx]
    render_uid = f"{mode}_{trial['trial_id']}_{idx}"

    if st.session_state["current_render_id"] != render_uid:
        st.session_state["current_render_id"] = render_uid
        st.session_state["trial_start_ts"] = time.time()
        st.session_state["trial_phase"] = "initial"
        st.session_state["initial_decision"] = None
        st.session_state["initial_rt_ms"] = None
        st.session_state["resolved_image_path"] = ""

    if mode == "formal" and not st.session_state.get("exp_start_ts"):
        st.session_state["exp_start_ts"] = time.time()

    if mode == "formal":
        elapsed = int(time.time() - st.session_state["exp_start_ts"])
        em, es = elapsed // 60, elapsed % 60
        st.markdown(
            f"<div style='text-align:center;padding:6px 0;font-size:0.95rem;'>"
            f"⏱️ 已用时 <b>{em:02d}:{es:02d}</b> &nbsp;|&nbsp; 题目进度 <b>{idx + 1} / {total}</b>"
            f"</div>",
            unsafe_allow_html=True,
        )
        st.progress((idx + 1) / total)
    else:
        st.progress((idx + 1) / total, text=f"练习题进度：{idx + 1}/{total}")

    st.subheader(f"{'练习题' if mode == 'practice' else '正式题'} {idx + 1} / {total}")

    phase = st.session_state.get("trial_phase", "initial")
    c1, c2 = st.columns([1.2, 1.0])

    with c1:
        try:
            img_path = resolve_image_path(trial["image_path"])
            st.session_state["resolved_image_path"] = str(img_path)
            st.image(Image.open(img_path).convert("RGB"), use_container_width=True)
            if st.session_state.get("show_debug"):
                st.caption(f"图片路径：{img_path}")
        except Exception as e:
            st.error(f"图片读取失败：{e}")

    with c2:
        if phase == "initial":
            st.markdown("### 第一步：请先独立判断")
            st.info("请仔细观察图像，在看到 AI 建议之前，先给出你的初步判断。")
            st.markdown("---")
            st.markdown("**你的初步判断：**")
            col_ok, col_ng = st.columns(2)

            def submit_initial(decision: str):
                rt = int((time.time() - st.session_state["trial_start_ts"]) * 1000)
                st.session_state["initial_decision"] = decision
                st.session_state["initial_rt_ms"] = rt
                st.session_state["trial_phase"] = "final"
                st.session_state["trial_start_ts"] = time.time()
                st.rerun()

            with col_ok:
                if st.button("✅ OK", key=f"i_ok_{render_uid}", use_container_width=True):
                    submit_initial("OK")
            with col_ng:
                if st.button("❌ NG", key=f"i_ng_{render_uid}", use_container_width=True):
                    submit_initial("NG")

            st.caption("完成初步判断后，将显示 AI 建议，再进行最终决策。" if mode != "practice" else "练习题不计入正式数据。")

        else:
            st.markdown("### 第二步：参考 AI 建议，做最终判断")
            st.info(f"**AI 判定：{normalize_label(trial['ai_suggestion'])}**")
            if trial["explanation_mode"] not in ("无解释", "") and safe_str(trial["explanation_text"]):
                st.write(trial["explanation_text"])
            st.caption(f"你的初步判断：**{st.session_state['initial_decision']}**")
            st.markdown("---")
            st.markdown("**你的最终判断：**")
            col_ok, col_ng = st.columns(2)

            def submit_final(decision: str):
                rt = int((time.time() - st.session_state["trial_start_ts"]) * 1000)
                init_dec = st.session_state["initial_decision"]
                init_rt = st.session_state["initial_rt_ms"] or 0
                meta = st.session_state["participant_meta"]
                exp_meta = st.session_state["exp_meta"]

                explanation_map = {
                    "无解释": "none",
                    "有解释": "unified",
                    "指标型解释": "metric",
                    "理由型解释": "reason",
                }

                record = {
                    "participant_id": meta.get("participant_id", ""),
                    "exp_type": meta.get("exp_type", ""),
                    "exp_condition": meta.get("exp_condition", ""),
                    "design": exp_meta.get("design", ""),
                    "trial_index": idx + 1,
                    "trial_id": trial["trial_id"],
                    "item_id": trial["item_id"],
                    "category": normalize_category(trial["category"]),
                    "defect_type": safe_str(trial["defect_type"]),
                    "complexity": normalize_complexity(trial["complexity"]),
                    "true_label": normalize_label(trial["true_label"]),
                    "ai_label": normalize_label(trial["ai_suggestion"]),
                    "ai_correct": parse_ai_correct(trial["ai_correct"], trial["ai_suggestion"], trial["true_label"]),
                    "explanation_mode": explanation_map.get(trial["explanation_mode"], safe_str(trial["explanation_mode"])),
                    "initial_decision": normalize_label(init_dec),
                    "final_decision": normalize_label(decision),
                    "initial_correct": decision_is_correct(init_dec, trial["true_label"]),
                    "final_correct": decision_is_correct(decision, trial["true_label"]),
                    "initial_rt_ms": init_rt,
                    "final_rt_ms": rt,
                    "total_rt_ms": init_rt + rt,
                    "decision_changed": int(normalize_label(decision) != normalize_label(init_dec)),
                    "adopt_ai": calc_adoption(decision, trial["ai_suggestion"]),
                    "dependence_type": calc_dependence_code(decision, trial["ai_suggestion"], trial["ai_correct"], trial["true_label"]),
                    "trial_stage": mode,
                    "recorded_at": datetime.now().isoformat(timespec="seconds"),
                }
                if mode != "practice":
                    st.session_state["responses"].append(record)
                    save_progress()
                st.session_state["current_index"] += 1
                st.session_state["current_render_id"] = None
                st.session_state["trial_phase"] = "initial"
                st.rerun()

            with col_ok:
                if st.button("✅ OK", key=f"f_ok_{render_uid}", use_container_width=True):
                    submit_final("OK")
            with col_ng:
                if st.button("❌ NG", key=f"f_ng_{render_uid}", use_container_width=True):
                    submit_final("NG")

            st.caption("练习题不计入正式数据。" if mode == "practice" else "点击后自动进入下一题。")


def render_questionnaire():
    st.title("实验结束问卷")
    st.markdown("请根据你在实验中的真实感受作答，没有对错之分。")
    exp_name = st.session_state["exp_meta"].get("exp_name", "")

    with st.form("questionnaire_form"):
        st.markdown("### 第一部分：对 AI 系统的整体评价")
        understanding = st.slider(
            "1. 我能够理解 AI 给出该判断的依据。",
            1, 7, 4,
            help="1=完全不理解，7=完全理解"
        )
        trust = st.slider(
            "2. 我认为该 AI 系统的判断总体上值得信任。",
            1, 7, 4,
            help="1=完全不信任，7=完全信任"
        )
        reliance = st.slider(
            "3. 在做最终判断时，我在多大程度上参考了 AI 的建议？",
            1, 7, 4,
            help="1=完全没有参考，7=完全依照AI建议"
        )
        ai_helpfulness = st.slider(
            "4. AI 的建议对我完成判断任务有帮助。",
            1, 7, 4,
            help="1=完全没帮助，7=非常有帮助"
        )

        st.markdown("### 第二部分：认知负荷评估（NASA-TLX）")
        nasa_mental = st.slider("5. 脑力需求：完成任务需要多少脑力投入？", 0, 100, 50, help="0=非常低，100=非常高")
        nasa_temporal = st.slider("6. 时间压力：你感受到多大的时间压力？", 0, 100, 50, help="0=非常低，100=非常高")
        nasa_effort = st.slider("7. 努力程度：你需要付出多少努力来完成任务？", 0, 100, 50, help="0=非常低，100=非常高")
        nasa_frustration = st.slider("8. 挫败感：你在任务中感到多少挫败、烦躁或压力？", 0, 100, 50, help="0=非常低，100=非常高")
        nasa_performance = st.slider("9. 你对自己在任务中表现的满意程度如何？", 0, 100, 50, help="0=非常不满意，100=非常满意")

        st.markdown("### 第三部分：判断策略")
        strategy = st.radio(
            "10. 在做最终判断时，你通常的策略是？",
            ["主要依靠自己的图像判断", "图像判断和AI建议各参考一半", "主要参考AI建议", "视情况而定"],
            index=3
        )
        changed_reason = st.radio(
            "11. 当你改变了初步判断时，主要原因是？",
            ["AI建议与我不同，选择相信AI", "AI的解释让我重新审视图像", "不确定时偏向跟随AI", "我没有改变过判断"],
            index=3
        )

        extra = {}
        if exp_name == "实验二":
            st.markdown("### 第四部分：解释形式对比（实验二专属）")
            easier = st.radio(
                "12. 你觉得哪种解释形式更容易理解？",
                ["指标型解释（数字/百分比）", "理由型解释（文字描述）", "两者差不多"]
            )
            more_trust = st.radio(
                "13. 你觉得哪种解释形式更让你信任AI的判断？",
                ["指标型解释（数字/百分比）", "理由型解释（文字描述）", "两者差不多"]
            )
            more_helpful = st.radio(
                "14. 你觉得哪种解释形式对你的最终判断帮助更大？",
                ["指标型解释（数字/百分比）", "理由型解释（文字描述）", "两者差不多"]
            )
            extra = {
                "easier_type": map_exp2_choice_code(easier),
                "trust_type": map_exp2_choice_code(more_trust),
                "helpful_type": map_exp2_choice_code(more_helpful),
            }

        comments = st.text_area(
            "15. 如有其他想说的（例如：哪些题目较难、对实验的建议等），请在此填写：",
            placeholder="选填"
        )

        submitted = st.form_submit_button("提交问卷", type="primary", use_container_width=True)

    if submitted:
        st.session_state["questionnaire"] = {
            "understanding": understanding,
            "trust": trust,
            "reliance": reliance,
            "ai_helpfulness": ai_helpfulness,
            "nasa_mental": nasa_mental,
            "nasa_temporal": nasa_temporal,
            "nasa_effort": nasa_effort,
            "nasa_frustration": nasa_frustration,
            "nasa_performance": nasa_performance,
            "strategy": map_strategy_code(strategy),
            "changed_reason": map_changed_reason_code(changed_reason),
            "comments": safe_str(comments),
            **extra,
        }
        save_progress()
        st.session_state["stage"] = "finish"
        st.rerun()


def render_finish():
    st.title("🎉 实验完成")
    st.success("感谢你的参与，数据已自动保存。")
    participant_id = st.session_state["participant_meta"].get("participant_id", "unknown")
    out_dir = ensure_results_dir(participant_id)
    st.markdown("研究者可在以下位置找到保存文件：")
    st.code(str(out_dir / OUTPUT_XLSX_NAME), language="text")

    if st.button("🏠 下一位被试（返回首页）", use_container_width=True, type="primary"):
        reset_experiment()
        st.session_state["stage"] = "setup"
        st.rerun()


def render_sidebar():
    with st.sidebar:
        st.title("管理员设置")
        st.session_state["show_debug"] = st.checkbox("显示调试信息", value=False)
        if st.session_state.get("exp_meta"):
            with st.expander("当前会话信息"):
                st.write(st.session_state.get("exp_meta", {}))
                st.write(st.session_state.get("participant_meta", {}))
        st.markdown("---")
        if st.button("重置当前会话", use_container_width=True):
            reset_experiment()
            st.rerun()


def validate_ready(df: pd.DataFrame) -> bool:
    if df.empty:
        st.warning("尚未加载题库。")
        return False
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"题库缺少必要列：{missing}")
        return False
    if not Path(DATASET_ROOT).exists():
        st.warning(f"当前图片根目录不存在：{DATASET_ROOT}。如果图片不显示，请优先检查该路径。")
    return True


def main():
    init_session()
    render_sidebar()
    df = read_bank()
    stage = st.session_state["stage"]

    if stage == "setup":
        render_setup(df)
    else:
        if not validate_ready(df):
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
