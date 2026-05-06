# urdf2excel

从 URDF 文件中提取**关节限位**和**连杆惯量**信息，一键输出到格式化 Excel 表格。

> Extracts joint limits & link inertial data from URDF files → formatted Excel (.xlsx).

---

## 功能 Features

- 递归扫描 `robots/` 目录，自动识别所有 `.urdf` 文件
- 也可直接指定任意 URDF 文件路径（支持绝对/相对路径）
- **Sheet 1 — 关节限位**：弧度值 + 角度值双列，含力矩/速度上限
- **Sheet 2 — 连杆惯量**：质量、重心坐标、6 个惯量张量分量
- 仅依赖 `openpyxl`，无需 ROS / MuJoCo 等机器人框架

---

## 安装 Installation

> 支持 Windows / Ubuntu / macOS，推荐使用 conda 管理 Python 环境。

### 0. 安装 conda（首次，仅需一次）

<details>
<summary><b>Windows</b></summary>

1. 前往 https://docs.conda.io/en/latest/miniconda.html  
   下载 **Miniconda3 Windows 64-bit**（Python 3.11）并双击安装。

2. 安装完成后打开 **开始菜单 → Anaconda Prompt**。

3. 验证安装：
```bat
conda --version
```

</details>

<details>
<summary><b>Ubuntu / Linux</b></summary>

```bash
# 下载安装脚本（x86_64，如为 ARM 请替换对应版本）
wget https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh
bash Miniconda3-latest-Linux-x86_64.sh
# 按提示操作，选择 yes 初始化 shell
# 完成后重开终端，或执行：
source ~/.bashrc

# 验证
conda --version
```

</details>

<details>
<summary><b>macOS</b></summary>

```bash
# Apple Silicon (M1/M2/M3)
curl -O https://repo.anaconda.com/miniconda/Miniconda3-latest-MacOSX-arm64.sh
bash Miniconda3-latest-MacOSX-arm64.sh

# Intel Mac
# curl -O https://repo.anaconda.com/miniconda/Miniconda3-latest-MacOSX-x86_64.sh
# bash Miniconda3-latest-MacOSX-x86_64.sh

# 按提示操作，完成后重开终端
conda --version
```

</details>

---

### 1. 创建项目环境（只需一次）

```bash
# 进入项目目录
cd /path/to/urdf2excel

# 创建 conda 环境（约 1 分钟）
conda env create -f environment.yml
```

### 2. 激活环境

```bash
# macOS / Linux / Windows Anaconda Prompt
conda activate urdf2excel
```

> Windows PowerShell 首次使用需先运行一次：
> ```powershell
> conda init powershell
> # 重开 PowerShell 后再 activate
> ```

---

## 使用 Usage

```bash
# ★ 推荐：指定单个 URDF 文件路径
python urdf2excel.py robots/lumos/nix2/urdf/nix2_joint21.urdf

# 使用绝对路径
python urdf2excel.py /absolute/path/to/robot.urdf

# 指定输出文件名
python urdf2excel.py path/to/robot.urdf -o output/my_robot.xlsx

# 递归扫描整个 robots/ 目录（每个 URDF 单独生成一个 xlsx）
python urdf2excel.py --robots robots/

# 扫描指定子目录
python urdf2excel.py --robots robots/lumos/

# 只列出找到的 URDF 文件，不生成 Excel
python urdf2excel.py --list
```

输出文件默认保存到 `output/<robot_stem>_<timestamp>.xlsx`。

---

## Excel 输出格式

### Sheet 1 — 关节限位

| 关节名称 | 弧度值 (rad) | | 角度值 (°) | | 力矩 (N·m) | 速度 (rad/s) | 类型 |
|----------|------------|---|----------|---|------------|-------------|------|
| | 最大值 | 最小值 | 最大值 | 最小值 | | | |
| torso_joint | 1.5708 | -1.5708 | 90.0 | -90.0 | 102 | 14.6 | revolute |
| left_hip_pitch_joint | 1.5708 | -1.6755 | 90.0 | -96.0 | 102 | 14.6 | revolute |

### Sheet 2 — 连杆惯量

| 连杆名称 | 质量 (kg) | 重心 X (m) | 重心 Y (m) | 重心 Z (m) | Ixx | Ixy | Ixz | Iyy | Iyz | Izz |
|----------|-----------|-----------|-----------|-----------|-----|-----|-----|-----|-----|-----|
| pelvis | 1.680900 | -0.001300 | 0.000000 | -0.000900 | 0.003073 | … | | | | |

---

## 目录结构

```
urdf2excel/
├── urdf2excel.py       ← 主脚本（单文件，无额外依赖）
├── environment.yml     ← conda 环境配置
├── .gitignore          ← 忽略 output/ 和 robots/
├── robots/             ← 放置机器人 URDF（不上传 git）
│   └── <robot>/
│       └── urdf/
│           └── *.urdf
└── output/             ← Excel 输出目录（不上传 git，自动创建）
```

---

## 依赖

| 包 | 版本 | 说明 |
|----|------|------|
| Python | ≥ 3.9 | 标准库：`xml`、`math`、`argparse`、`pathlib` |
| openpyxl | ≥ 3.1 | Excel 读写 |

---

## 常见问题

**Q: 提示 `No .urdf files found`？**  
A: 确认 URDF 文件在 `robots/` 目录下，或使用绝对路径：`python urdf2excel.py /full/path/to/robot.urdf`

**Q: Windows 路径含空格？**  
A: 用引号包裹：`python urdf2excel.py "C:\My Robots\robot.urdf"`

**Q: 更新 conda 环境？**
```bash
conda env update -f environment.yml --prune
```
