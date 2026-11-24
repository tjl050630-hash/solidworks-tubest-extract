# SolidWorks → TubesT Auto Tube Extraction Skill (`solidworks-tubest-extract`)

一个面向**激光切管（柏楚 TubesT）**场景的 **Antigravity Skill + SolidWorks 自动化拆管脚本**。

只需提供机柜、机架或焊件的 `.SLDPRT`（多实体焊件）、`.SLDASM`（装配体）或 `.STEP` 模型，脚本即可自动调用本机 SolidWorks：
1. **自动剔除非管材实体**：过滤加强筋（Gusset）、封板、地脚板、门板、螺丝/螺母/合页等标准件；
2. **自动去重并统计数量**：焊件按切割清单（Cut-List）合并同类管件，装配体按零件路径+配置合并统计根数；
3. **双格式输出（适配柏楚 TubesT）**：
   - 批量导出**单根独立 `.step`**（按公制毫米尺寸与数量命名，如 `Frame_T01_50.8x76.2_L1276mm_2pcs.step`）；
   - 导出整机纯管合集 **`*_AllTubesOnly.sat`**（方便直接整框拖入 TubesT 自动拆管套料）；
   - 生成带 UTF-8 BOM 的 **`*_BOM.csv`**（Excel 打开不乱码）；
   - 自动保存原模型 vs 纯管材等轴测对比截图（`*_00_Original.png` / `*_00_TubesOnly.png`）；
4. **单实体强制校验 + 进程清理**：自动核查每个导出 `.step` 严格只含 `1` 个 `MANIFOLD_SOLID_BREP` 实体，并在完成后关闭脚本拉起的 SolidWorks 进程释放内存。

---

## 目录结构

```text
solidworks-tubest-extract/
├── SKILL.md                  # Antigravity / AI Agent 技能说明文件
├── scripts/
│   └── extract_tubes.ps1     # SolidWorks C# Interop 自动化拆管脚本
├── README.md
└── LICENSE
```

---

## 环境要求

- **操作系统**：Windows 10 / 11（带 Windows PowerShell 5.1+）
- **CAD 软件**：已安装本机版 **SolidWorks**（支持自动搜索 `C:\Program Files`、`D:\Program Files` 及注册表中的安装路径，也可手动通过 `-SolidWorksDir` 指定）
- **套料软件**：柏楚 **TubesT**（将导出的 `.step` 或 `.sat` 直接拖入使用）

---

## 安装与使用

### 方式一：作为 Antigravity 全局 Skill 安装

将整个 `solidworks-tubest-extract` 文件夹复制到本机全局技能目录：

```powershell
# Windows 全局 Skill 目录
%USERPROFILE%\.gemini\config\skills\solidworks-tubest-extract
```

之后在任意对话中直接发送指令即可触发，例如：
> “帮我把这个机柜模型拆管：`D:\Orders\Cabinet.SLDPRT`”

默认会把拆好的文件输出到桌面 `拆图\<项目名>\` 目录。

### 方式二：直接用 PowerShell 命令行运行

```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\extract_tubes.ps1" `
  -InputPath "D:\Models\Cabinet.SLDPRT" `
  -ProjectName "测试机柜"
```

#### 可选参数

| 参数 | 说明 | 默认值 |
| :--- | :--- | :--- |
| `-InputPath` | 待拆管模型路径（支持 `.SLDPRT` / `.SLDASM` / `.STEP` / `.STP`） | 必填 |
| `-ProjectName` | 输出子文件夹与文件前缀名 | 原文件名 |
| `-OutputRoot` | 输出根目录 | `%USERPROFILE%\Desktop\拆图` |
| `-SolidWorksDir` | SolidWorks 安装目录（含 Interop DLL） | 自动检测 |
| `-MinTubeCrossSectionMm` | 管材最小截面阈值（mm，用于过滤薄板/加强筋） | `12.0` |
| `-MinTubeLengthMm` | 管材最小长度阈值（mm，用于过滤短小连接件） | `50.0` |
| `-KeepSolidWorksOpen` | 运行完成后不关闭脚本拉起的 SolidWorks 进程 | 关闭 |

---

## 为什么不用 PowerShell 直接 `New-Object -ComObject`？

在 Windows PowerShell 中直接调用 `New-Object -ComObject SldWorks.Application` 容易触发 `0x8002802B (TYPE_E_ELEMENTNOTFOUND)` 晚绑定错误。本脚本通过 `[System.Reflection.Assembly]::LoadFrom` 预加载 SolidWorks 官方 Interop DLL（`SolidWorks.Interop.sldworks.dll`、`SolidWorks.Interop.swconst.dll`），再通过内嵌 C# 强类型调用 API，稳定性和速度远高于传统脚本。

---

## License

MIT License
