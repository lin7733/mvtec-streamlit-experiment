import hashlib
import json
import os
import random
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="MVTec AD 解释透明度实验",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "MVTec AD 解释透明度实验平台"
RESULTS_DIR_DEFAULT = "results"
PRACTICE_TRIALS = 4
BREAK_AFTER = 24
CONSENT_TEXT = """
我已知晓：本实验用于毕业论文研究，实验过程将记录我的作答、反应时及主观评分。
我理解实验过程中可随时退出，研究数据仅用于学术分析。
"""


def stable_hash_int(text: str) -> int:
    return int(hashlib.md5(text.encode("utf-8")).hexdigest(), 16)


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        df[c] = df[c].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    return df


def read_bank(uploaded_file, workbook_path: str) -> pd.DataFrame:
    if uploaded_file is not None:
        return normalize_df(pd.read_excel(uploaded_file, sheet_name="题库总表"))
    if workbook_path:
        return normalize_df(pd.read_excel(workbook_path, sheet_name="题库总表"))
    return pd.DataFrame()


def parse_ai_correct(v: str) -> bool:
    return str(v).startswith("正确")


def calc_adoption(decision: str, ai_suggestion: str) -> str:
    return "采纳" if decision == ai_suggestion else "未采纳"


def calc_dependence(decision: str, ai_suggestion: str, ai_correct: str) -> str:
    adopted = decision == ai_suggestion
    if parse_ai_correct(ai_correct):
        return "适当依赖" if adopted else "依赖不足"
    return "过度依赖" if adopted else "适当依赖"


def resolve_image_path(raw_path: str, dataset_root: str) -> Path:
    if not raw_path:
        raise FileNotFoundError("题库中未提供图片路径。")

    raw_path_norm = str(raw_path).replace("\\", "/")
    root = Path(dataset_root)  # 部署后这里是 Path("00_raw")

    # 1. 先试直接路径（本地绝对路径，部署后不会命中）
    direct_path = Path(raw_path)
    if direct_path.exists():
        return direct_path

    # 2. 从绝对路径中提取 00_raw 之后的部分（核心逻辑，不变）
    if "00_raw/" in raw_path_norm:
        rel = raw_path_norm.split("00_raw/", 1)[1]
        # root = Path("00_raw") 时，root.name == "00_raw"，走第一个分支
        candidate = root / Path(rel) if root.name == "00_raw" else root / "00_raw" / Path(rel)
        if candidate.exists():
            return candidate

    # 3. 兜底：按类别名截取（不变）
    parts = raw_path_norm.split("/")
    for key in ["bottle", "capsule", "metal_nut"]:
        if key in parts:
            idx = parts.index(key)
            rel = Path(*parts[idx:])
            candidate = root / rel if root.name == "00_raw" else root / "00_raw" / rel
            if candidate.exists():
                return candidate

    raise FileNotFoundError(f"无法解析图片路径：{raw_path}")


def ensure_results_dir(base_dir: str, participant_id: str) -> Path:
    out_dir = Path(base_dir) / participant_id
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def save_json(path: Path, payload: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def save_csv(path: Path, rows: list[dict]):
    if rows:
        pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")


def get_condition_from_participant(participant_id: str, options: list[str]) -> str:
    return options[stable_hash_int(participant_id) % len(options)]


def build_exp1_trials(df: pd.DataFrame, participant_id: str, manual_condition: str | None = None):
    condition = manual_condition or get_condition_from_participant(participant_id, ["无解释", "有解释"])
    trials = []
    for _, row in df.iterrows():
        explanation_text = row["实验一-无解释呈现"] if condition == "无解释" else row["实验一-统一解释内容"]
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
    rnd = random.Random(stable_hash_int(participant_id + "_exp1_order"))
    rnd.shuffle(trials)
    return trials, {"exp_name": "实验一", "condition": condition, "design": "组间"}


def build_exp2_trials(df: pd.DataFrame, participant_id: str):
    version = "A" if stable_hash_int(participant_id) % 2 == 0 else "B"
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
    rnd = random.Random(stable_hash_int(participant_id + "_exp2_order"))
    rnd.shuffle(trials)
    return trials, {"exp_name": "实验二", "counterbalance_version": version, "design": "被试内"}


def select_practice_trials(trials: list[dict], participant_id: str, n: int):
    if n <= 0:
        return []
    rnd = random.Random(stable_hash_int(participant_id + "_practice"))
    copied = trials.copy()
    rnd.shuffle(copied)
    practice = copied[: min(n, len(copied))]
    for t in practice:
        t["is_practice"] = True
    return practice


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
        "responses": [],
        "questionnaire": {},
        "finished": False,
        "rest_done": False,
        "show_debug": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_experiment():
    for k in [
        "stage", "participant_meta", "exp_meta", "trials", "practice_trials", "current_index",
        "current_render_id", "trial_start_ts", "responses", "questionnaire", "finished",
        "rest_done", "show_debug"
    ]:
        if k in st.session_state:
            del st.session_state[k]
    init_session()


def save_progress(results_dir: str):
    participant_id = st.session_state["participant_meta"].get("participant_id", "unknown")
    out_dir = ensure_results_dir(results_dir, participant_id)
    save_json(out_dir / "session_meta.json", {
        "participant_meta": st.session_state["participant_meta"],
        "exp_meta": st.session_state["exp_meta"],
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    })
    save_csv(out_dir / "trial_responses.csv", st.session_state["responses"])
    if st.session_state.get("questionnaire"):
        save_json(out_dir / "questionnaire.json", st.session_state["questionnaire"])


def render_setup(df: pd.DataFrame):
    st.title(APP_TITLE)
    st.write("请先在左侧设置题库路径、图片根目录和结果保存目录，然后在此页面输入被试信息并初始化实验。")
    st.info("提示：实验一是组间设计，同一位被试全程只会看到“无解释”或“有解释”中的一种，不会在一套题里混合出现。")

    with st.form("setup_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            participant_id = st.text_input("被试编号", placeholder="例如：P001")
            age = st.text_input("年龄（可选）")
        with c2:
            gender = st.selectbox("性别（可选）", ["", "女", "男", "其他/不便透露"])
            major = st.text_input("专业（可选）")
        with c3:
            exp_name = st.selectbox("实验类型", ["实验一：有无解释", "实验二：解释形式"])
            exp1_manual = st.selectbox("实验一条件（可手动指定）", ["自动分配", "无解释", "有解释"])

        submitted = st.form_submit_button("初始化实验", use_container_width=True)

    if submitted:
        if df.empty:
            st.error("尚未成功读取题库，请先在左侧上传或指定修订版 xlsx。")
            return
        if not participant_id:
            st.error("请先输入被试编号。")
            return

        st.session_state["participant_meta"] = {
            "participant_id": participant_id.strip(),
            "age": age.strip(),
            "gender": gender,
            "major": major.strip(),
        }

        if exp_name.startswith("实验一"):
            trials, meta = build_exp1_trials(df, participant_id.strip(), None if exp1_manual == "自动分配" else exp1_manual)
        else:
            trials, meta = build_exp2_trials(df, participant_id.strip())

        st.session_state["exp_meta"] = meta
        st.session_state["trials"] = trials
        st.session_state["practice_trials"] = select_practice_trials(trials, participant_id.strip(), PRACTICE_TRIALS)
        st.session_state["current_index"] = 0
        st.session_state["responses"] = []
        st.session_state["stage"] = "consent"
        st.rerun()

    with st.expander("查看当前题库摘要", expanded=False):
        if not df.empty:
            st.dataframe(df[["题号", "图片ID", "产品类别", "复杂度", "真实标签", "AI建议", "AI是否正确"]].head(12), hide_index=True, use_container_width=True)
            st.write(f"题目总数：{len(df)}")
            st.dataframe(df.groupby(["产品类别", "复杂度", "真实标签"]).size().reset_index(name="数量"), hide_index=True, use_container_width=True)


def render_consent():
    st.title("知情同意")
    st.info(CONSENT_TEXT)
    agree = st.checkbox("我已阅读并同意参加本实验。")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("返回设置页", use_container_width=True):
            st.session_state["stage"] = "setup"
            st.rerun()
    with c2:
        if st.button("进入说明页", type="primary", use_container_width=True, disabled=not agree):
            st.session_state["stage"] = "instruction"
            st.rerun()


def render_instruction():
    st.title("实验说明")
    exp_name = st.session_state["exp_meta"].get("exp_name", "")
    st.markdown(f"**当前实验：{exp_name}**")
    if exp_name == "实验一":
        condition = st.session_state["exp_meta"].get("condition", "")
        st.markdown(f"- 当前被试所属条件：**{condition}**")
        st.markdown("""
        - 你的任务是依据产品图像以及 AI 给出的建议，判断该产品是 **OK（合格）** 还是 **NG（不合格）**。
        - 在本实验中，你会始终看到同一种解释条件：要么只有 AI 判定，要么看到 AI 判定与统一解释。
        - 请尽量独立思考，不要机械跟随 AI。
        """)
    else:
        version = st.session_state["exp_meta"].get("counterbalance_version", "")
        st.markdown(f"- 当前被试的平衡版本：**{version}**")
        st.markdown("""
        - 你的任务是依据产品图像以及 AI 给出的建议，判断该产品是 **OK（合格）** 还是 **NG（不合格）**。
        - 本实验中你会看到两种解释形式：**指标型解释** 与 **理由型解释**。
        - 两种解释都会出现，请尽量根据图像与解释信息综合判断。
        """)

    st.markdown("""
    **操作方式**
    1. 先观察图像；
    2. 阅读 AI 建议与解释信息；
    3. 点击“判定为 OK（合格）”或“判定为 NG（不合格）”；
    4. 系统将自动记录反应时间。

    **注意**
    - 正式实验共 48 题。
    - 中途会安排一次短暂休息。
    - 结束后请填写理解、信任和认知负荷问卷。
    """)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("返回知情同意", use_container_width=True):
            st.session_state["stage"] = "consent"
            st.rerun()
    with c2:
        if st.button("进入练习题" if PRACTICE_TRIALS > 0 else "开始正式实验", type="primary", use_container_width=True):
            st.session_state["stage"] = "practice" if PRACTICE_TRIALS > 0 else "formal"
            st.session_state["current_index"] = 0
            st.session_state["current_render_id"] = None
            st.rerun()


def render_rest():
    st.title("请稍作休息")
    st.write("你已完成一半题目。建议稍作休息后继续。")
    if st.button("继续实验", type="primary", use_container_width=True):
        st.session_state["rest_done"] = True
        st.session_state["stage"] = "formal"
        st.session_state["current_render_id"] = None
        st.rerun()


def render_trial(trials: list[dict], mode: str, dataset_root: str, results_dir: str, show_debug: bool = False):
    idx = st.session_state["current_index"]
    total = len(trials)
    if idx >= total:
        st.session_state["current_index"] = 0
        st.session_state["stage"] = "formal" if mode == "practice" else "questionnaire"
        st.session_state["current_render_id"] = None
        st.rerun()

    trial = trials[idx]
    render_uid = f"{mode}_{trial['trial_id']}_{idx}"
    if st.session_state["current_render_id"] != render_uid:
        st.session_state["current_render_id"] = render_uid
        st.session_state["trial_start_ts"] = time.time()

    if mode == "formal" and idx == BREAK_AFTER and not st.session_state.get("rest_done", False):
        st.session_state["stage"] = "rest"
        st.rerun()

    st.progress(idx / total, text=f"{'练习题' if mode == 'practice' else '正式实验'}进度：{idx}/{total}")
    st.subheader(f"{'练习题' if mode == 'practice' else '正式题'} {idx + 1} / {total}")

    c1, c2 = st.columns([1.15, 1.0])
    with c1:
        if show_debug:
            st.caption(f"题号：{trial['trial_id']}｜图片ID：{trial['item_id']}｜类别：{trial['category']}｜复杂度：{trial['complexity']}")
        try:
            img_path = resolve_image_path(trial["image_path"], dataset_root)
            st.image(Image.open(img_path), use_container_width=True)
            if show_debug:
                st.caption(f"图片路径：{img_path}")
        except Exception as e:
            st.error(f"图片读取失败：{e}")

    with c2:
        st.markdown("### AI 输出")
        st.info(f"AI 判定：**{trial['ai_suggestion']}**")
        if trial["explanation_mode"] == "无解释":
            st.write("本条件下不提供额外解释信息。")
        else:
            if show_debug:
                st.caption(f"当前解释模式：{trial['explanation_mode']}")
            st.write(trial["explanation_text"])

        st.markdown("---")
        st.markdown("### 你的判断")
        c_ok, c_ng = st.columns(2)

        def submit_response(decision: str):
            rt_ms = int((time.time() - st.session_state["trial_start_ts"]) * 1000)
            record = {
                "participant_id": st.session_state["participant_meta"].get("participant_id", ""),
                "exp_name": trial["exp_name"],
                "trial_stage": mode,
                "trial_index": idx + 1,
                "trial_id": trial["trial_id"],
                "item_id": trial["item_id"],
                "category": trial["category"],
                "defect_type": trial["defect_type"],
                "complexity": trial["complexity"],
                "true_label": trial["true_label"],
                "ai_suggestion": trial["ai_suggestion"],
                "ai_correct": trial["ai_correct"],
                "explanation_mode": trial["explanation_mode"],
                "decision": decision,
                "adoption": calc_adoption(decision, trial["ai_suggestion"]),
                "dependence_type": calc_dependence(decision, trial["ai_suggestion"], trial["ai_correct"]),
                "rt_ms": rt_ms,
                "timestamp": datetime.now().isoformat(timespec="seconds"),
            }
            if mode != "practice":
                st.session_state["responses"].append(record)
                save_progress(results_dir)
            st.session_state["current_index"] += 1
            st.session_state["current_render_id"] = None
            st.rerun()

        with c_ok:
            if st.button("判定为 OK（合格）", key=f"ok_{render_uid}", use_container_width=True):
                submit_response("OK（合格）")
        with c_ng:
            if st.button("判定为 NG（不合格）", key=f"ng_{render_uid}", use_container_width=True):
                submit_response("NG（不合格）")

        st.caption("练习题不计入正式数据。" if mode == "practice" else "请尽量快速且认真地作答。点击按钮后将自动进入下一题。")


def render_questionnaire(results_dir: str):
    st.title("实验结束问卷")
    exp_name = st.session_state["exp_meta"].get("exp_name", "")
    with st.form("questionnaire_form"):
        st.markdown("### 主观量表")
        understanding = st.slider("我能够理解 AI 给出该判断的依据。", 1, 7, 4)
        trust = st.slider("我认为该 AI 建议总体上值得信任。", 1, 7, 4)
        nasa = st.slider("完成本实验时，你感受到的总体认知负荷是多少？", 0, 100, 50)
        extra = {}
        if exp_name == "实验二":
            st.markdown("### 实验二补充问题")
            easier = st.radio("你觉得哪种解释更容易理解？", ["指标型解释", "理由型解释", "差不多"])
            more_trust = st.radio("你觉得哪种解释更值得信任？", ["指标型解释", "理由型解释", "差不多"])
            extra = {"easier_to_understand": easier, "more_trustworthy": more_trust}
        comments = st.text_area("补充意见（可选）", placeholder="例如：哪些题目较难、哪种解释更有帮助等")
        submitted = st.form_submit_button("提交问卷并结束实验", type="primary", use_container_width=True)

    if submitted:
        st.session_state["questionnaire"] = {
            "understanding": understanding,
            "trust": trust,
            "nasa_tlx": nasa,
            "comments": comments.strip(),
            **extra,
        }
        save_progress(results_dir)
        st.session_state["stage"] = "finish"
        st.rerun()


def render_finish(results_dir: str):
    st.title("实验完成 🎉")
    st.success("感谢参与！请下载并发送你的数据文件给研究者。")

    responses = pd.DataFrame(st.session_state["responses"])
    participant_id = st.session_state["participant_meta"].get("participant_id", "unknown")

    if not responses.empty:
        # 显示简要统计（原有代码保留）
        ...

        # ✅ 关键：让被试自己下载数据
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ 下载实验数据（请发给研究者）",
                data=responses.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"{participant_id}_trials.csv",
                mime="text/csv",
                use_container_width=True,
                type="primary",
            )
        with col2:
            q = st.session_state.get("questionnaire", {})
            if q:
                st.download_button(
                    "⬇️ 下载问卷数据",
                    data=pd.DataFrame([q]).to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"{participant_id}_questionnaire.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
    if st.button("开始下一位被试", use_container_width=True):
        reset_experiment()
        st.rerun()


def render_sidebar():
    st.sidebar.title("实验设置")
    uploaded_file = st.sidebar.file_uploader("上传修订版题库 xlsx", type=["xlsx"])

    # ❌ 原来：让用户手动输入路径
    # workbook_path = st.sidebar.text_input("或填写本地题库路径", value="")
    # dataset_root = st.sidebar.text_input("图片根目录（MVTec_AD_Thesis 或 00_raw）", value="")

    # ✅ 改为：硬编码相对路径（部署后固定不变）
    workbook_path = "05_metadata/MVTec_实验题库_完整版_解释优化版.xlsx"
    dataset_root = "00_raw"  # 相对于 app.py 的位置

    results_dir = RESULTS_DIR_DEFAULT  # "results"，直接用默认值

    # 调试开关仅管理员需要，可保留
    st.session_state["show_debug"] = False  # 部署后关闭调试
    ...
    return uploaded_file, workbook_path, dataset_root, results_dir


def validate_ready(df: pd.DataFrame, dataset_root: str):
    if df.empty:
        st.warning("尚未加载题库。")
        return False
    required_cols = [
        "题号", "图片ID", "产品类别", "复杂度", "真实标签", "AI建议", "AI是否正确",
        "实验一-无解释呈现", "实验一-统一解释内容", "实验二-指标型解释内容",
        "实验二-理由型解释内容", "图片源路径"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"题库缺少必要列：{missing}")
        return False
    if not dataset_root:
        st.warning("请先在左侧填写图片根目录。")
        return False
    return True


def main():
    init_session()
    uploaded_file, workbook_path, dataset_root, results_dir = render_sidebar()
    df = read_bank(uploaded_file, workbook_path)
    ready = validate_ready(df, dataset_root) if st.session_state["stage"] != "setup" else True
    stage = st.session_state["stage"]
    if stage == "setup":
        render_setup(df)
    elif not ready:
        st.stop()
    elif stage == "consent":
        render_consent()
    elif stage == "instruction":
        render_instruction()
    elif stage == "practice":
        render_trial(st.session_state["practice_trials"], "practice", dataset_root, results_dir, st.session_state["show_debug"])
    elif stage == "formal":
        render_trial(st.session_state["trials"], "formal", dataset_root, results_dir, st.session_state["show_debug"])
    elif stage == "rest":
        render_rest()
    elif stage == "questionnaire":
        render_questionnaire(results_dir)
    elif stage == "finish":
        render_finish(results_dir)


if __name__ == "__main__":
    main()
