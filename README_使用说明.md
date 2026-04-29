# MVTec AD 解释透明度实验（Streamlit）

这个项目用于运行你当前毕业论文设计下的两个实验：

- **实验一**：有无解释（组间）
- **实验二**：解释形式（指标型 vs 理由型，被试内，含 counterbalance）

题库基于你的修订版 Excel：
- `题库总表`
- `实验设计说明`
- `被试数据记录模板`

---

## 一、文件准备

你需要准备三类文件：

### 1）修订版题库 Excel
例如：
`MVTec_实验题库_完整版_修订版.xlsx`

### 2）MVTec 图片文件夹
建议保留你原来的目录，例如：

```text
MVTec_AD_Thesis/
└── 00_raw/
    ├── bottle/
    ├── capsule/
    └── metal_nut/
```

运行时，在 Streamlit 侧边栏填写：
- `MVTec_AD_Thesis` 路径
或
- 直接填写 `00_raw` 路径

都可以。

### 2.1）公网部署图片目录

如果部署到 Streamlit Cloud，推荐使用压缩后的图片目录：

```text
streamlit_mvtec_experiment/
└── 00_raw_compressed/
    ├── bottle/
    ├── capsule/
    └── metal_nut/
```

这个目录必须和 `app.py` 一起提交到 GitHub 当前部署分支。Streamlit Cloud 页面里的
`/mount/src/.../00_raw_compressed` 是云端容器路径；如果它提示不存在，通常说明
`00_raw_compressed` 没有被提交或没有推送到正在部署的仓库分支。

当前程序会优先读取 `00_raw_compressed`，找不到时再尝试 `00_raw`。如果你使用其他目录名，
可以在部署环境变量中设置：

```text
MVTEC_DATASET_ROOT=你的图片目录名
```

### 3）可选：解释文本优化脚本
如果你想根据 `ground_truth` 自动重写解释文本，可运行：

```text
根据ground_truth更新解释文本.py
```

运行前把脚本顶部的 `INPUT_XLSX` 和 `OUTPUT_XLSX` 改成你自己电脑里的路径。

---

## 二、安装依赖

```bash
pip install -r requirements.txt
```

---

## 三、启动方式

```powershell
.\start_app.ps1
```

如果你仍然想手动启动，请明确使用项目虚拟环境：

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

---

## 四、实验流程

### 实验一
- 被试只进入 **无解释** 或 **有解释** 其中一个条件
- 条件可自动分配，也可手动指定
- 完成全部 48 题
- **注意：实验一是组间设计，所以同一位被试不会在一套题里同时看到两种条件**

### 实验二
- 被试完成全部 48 题
- 每题只出现一次
- 系统在 `产品类别 × 复杂度 × 真实标签` 内尽量平衡：
  - 一半题为 **指标型解释**
  - 一半题为 **理由型解释**
- 不同被试采用 A/B counterbalance 版本

### 正式流程
1. 设置被试信息
2. 知情同意
3. 说明页
4. 练习题（默认 4 题，可在 `app.py` 中修改）
5. 正式实验（48 题）
6. 中途休息一次
7. 结束问卷
8. 保存结果

---

## 五、结果文件

默认保存到：

```text
results/
└── P001/
    ├── session_meta.json
    ├── trial_responses.csv
    └── questionnaire.json
```

---

## 六、CSV 核心字段说明

- `participant_id`：被试编号
- `exp_name`：实验一 / 实验二
- `trial_stage`：practice / formal
- `trial_id`：题号
- `item_id`：图片ID
- `complexity`：低复杂 / 高复杂
- `true_label`：真实标签
- `ai_suggestion`：AI建议
- `ai_correct`：AI是否正确
- `explanation_mode`：无解释 / 有解释 / 指标型解释 / 理由型解释
- `decision`：被试最终决策
- `adoption`：采纳 / 未采纳
- `dependence_type`：适当依赖 / 过度依赖 / 依赖不足
- `rt_ms`：反应时间（毫秒）

---

## 七、界面调整说明

新版 `app.py` 已做这些修改：
- 默认**隐藏**题号、图片ID、类别、复杂度、图片路径
- 默认**不显示**“解释模式”标签，避免额外提示被试
- 侧边栏新增：
  - `显示调试信息（题号/图片ID/路径/解释模式）`
- 侧边栏会显示当前会话信息，便于你核对被试分配到的实验条件

---

## 八、你可以按需要调整的地方

在 `app.py` 顶部有几个常量可以直接改：

- `PRACTICE_TRIALS = 4`：练习题数量
- `BREAK_AFTER = 24`：做到第几题时休息
- `RESULTS_DIR_DEFAULT = "results"`：默认保存目录

---

## 九、建议

### 1）正式实验前先做 2–5 人试跑
重点检查：
- 图片路径是否能正确读取
- 反应时间是否正常记录
- 指标型和理由型文本是否自然
- 练习题数量是否合适
- 48 题整体时长是否可接受

### 2）先用“有解释”条件单独试一次实验一
这样可以快速确认 `实验一-统一解释内容` 是否正常显示。

### 3）如果要基于 mask 重写解释
先运行：

```bash
python 根据ground_truth更新解释文本.py
```

然后把生成的新 Excel 作为 Streamlit 的题库输入。

---

## 十、提醒

由于你的 Excel 中保存的是 Windows 本地绝对路径，因此：
- **最好在你自己的电脑本地运行这个 Streamlit 应用**
- 并在侧边栏填写你本机的 `MVTec_AD_Thesis` 或 `00_raw` 路径
