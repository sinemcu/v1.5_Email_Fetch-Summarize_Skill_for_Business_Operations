---
name: summarise
description: "整合 WIP 附件为晟矽微运营汇总表格。可独立运行，或与 email-fetch 配合实现下载→整合一条龙。自动检测供应商、分配配色、生成29列标准格式。"
---

# Summarise — 运营 WIP 数据整合

将多封邮件下载的各供应商 WIP 附件整合为统一的**晟矽微运营汇总表格**（29 列标准格式，带供应商分组配色）。

---

## 触发条件

当用户要求：
- 整合邮件下载的 WIP 附件为汇总表
- 按运营格式输出/对齐运营汇总表
- 将下载的数据整理成运营表格
- 与 email-fetch 配合使用：下载指定日期的邮件附件后立即整合

---

## 前置技能

- **email-fetch** — 用于从 OWA 邮箱下载 WIP 附件（可选，也可手动放入附件目录）
- 依赖 Python 包：`openpyxl`、`xlrd`

---

## 工作目录约定

| 路径 | 用途 |
|------|------|
| `~/Downloads/owa_attachments/` | 邮件附件存放根目录 |
| `~/Downloads/owa_attachments/<YYYY-MM-DD>/` | 按日期存放的附件子目录 |
| `~/Downloads/owa_attachments/logs/` | 共享日志目录 |
| `~/Downloads/owa_attachments/logs/email-fetch-YYYY-MM-DD.log` | email-fetch 执行日志（可解析） |
| `~/Downloads/owa_attachments/logs/summarise-YYYY-MM-DD.log` | summarise 执行日志 |
| `~/Downloads/` | 最终输出目录 |
| `scripts/consolidate_wip.py` | 整合脚本（相对 SKILL.md 所在目录） |
| `scripts/validate_wip.py` | 校验脚本 |

---

## Workflow

```
Phase 0  确认附件 ── 确认日期子目录，可选读取 email-fetch 日志获取附件清单
Phase 1  运行整合 ── python3 scripts/consolidate_wip.py <日期子目录路径>
Phase 2  数据校验 ── python3 scripts/validate_wip.py
Phase 3  写入日志 ── 将本次执行记录写入共享日志目录
Phase 4  报告结果 ── 告知用户文件路径、行数、校验结果、日志引用
```

### Phase 0 — 确认附件

脚本默认从日期子目录读取附件。有两种方式获取文件清单：

#### 方式 A：从 email-fetch 日志读取（推荐，当 email-fetch 刚执行过时）

email-fetch 的日志格式为：
```
[YYYY-MM-DD HH:MM:SS] [SUCCESS] [DOWNLOAD] 下载完成: 文件名 (大小)
[YYYY-MM-DD HH:MM:SS] [ERROR] [DOWNLOAD] 下载失败: 文件名 — 原因
```

解析 email-fetch 日志提取文件清单：

```bash
DATE="2026-05-29"
EXEC_DATE="$(date '+%Y-%m-%d')"  # email-fetch 的执行日期（通常与邮件日期同一天）
LOG_FILE="$HOME/Downloads/owa_attachments/logs/email-fetch-${EXEC_DATE}.log"

if [ -f "${LOG_FILE}" ]; then
  # 提取成功下载的文件名
  echo "=== 从 email-fetch 日志解析附件清单 ==="
  grep '\[SUCCESS\] \[DOWNLOAD\] 下载完成:' "${LOG_FILE}" | \
    sed -E 's/.*下载完成: (.*) \(.*/\1/' | sort
  echo "=== 下载失败的文件 ==="
  grep '\[ERROR\] \[DOWNLOAD\] 下载失败:' "${LOG_FILE}" | \
    sed -E 's/.*下载失败: (.*) —.*/\1/' | sort
else
  echo "未找到 email-fetch 日志: ${LOG_FILE}，将直接从目录扫描文件"
fi
```

如果附件在日志中标记为**下载失败**，提醒用户该文件可能不在附件目录中。

#### 方式 B：直接扫描目录

```bash
DATE="2026-05-29"
TARGET_DIR="~/Downloads/owa_attachments/${DATE}"

ls -la "${TARGET_DIR}"
```

确认目录中包含 WIP 相关的 `.xlsx` / `.xls` 文件。

#### 如果用户未明确日期

先列出已有的日期子目录供用户选择：

```bash
ls -la ~/Downloads/owa_attachments/ | grep "^d"
```

#### 附件不在预期目录时

如果附件在其他位置，先拷贝到该日期子目录：

```bash
DATE="2026-05-29"
TARGET_DIR="~/Downloads/owa_attachments/${DATE}"
mkdir -p "${TARGET_DIR}"
cp <源文件路径> "${TARGET_DIR}/"
```

### Phase 1 — 运行整合脚本

```bash
python3 scripts/consolidate_wip.py ~/Downloads/owa_attachments/2026-05-29/
```

脚本自动完成：
1. 读取指定目录下所有支持的 WIP 文件
2. 将各供应商不同格式的源数据映射为 **29 列标准格式**
3. **自动检测供应商**，按数据量从大到小分配配色（已知供应商保持原配色）
4. 应用表头深蓝底白字、细边框、冻结窗格 A2、列宽 14
5. 输出后自动运行 `validate_wip.py` 校验数据准确性

### Phase 2 — 数据校验

运行校验脚本，自动检查数据准确性：

```bash
python3 scripts/validate_wip.py
```

校验项目包括：

| 检查项 | 说明 |
|--------|------|
| 输出文件可打开 | 文件存在且可正常读取 |
| 列结构 29 列 | 表头名称完全匹配 |
| 行数对比 | 源文件数据行总计 ≈ 输出行数 |
| 数值列检查 | 数值列不含文本异常值 |
| 重复行检查 | 警告同一供应商下客户代码+芯片型号重复 |
| 供应商覆盖 | 所有源文件的供应商都已纳入 |
| 空数据列报告 | 标注超过 80% 为空的列（通常正常） |
| 抽样检查 | 每个供应商随机抽样 1 行供人工核对 |

**退出码：** 0 = 全部通过，1 = 存在 FAIL 项

### Phase 3 — 写入日志

将本次执行信息写入共享日志目录：

```bash
EXEC_DATE="$(date '+%Y-%m-%d')"
DATE="2026-05-29"
LOG_DIR="$HOME/Downloads/owa_attachments/logs"
LOG_FILE="${LOG_DIR}/summarise-${EXEC_DATE}.log"
TARGET_DIR="~/Downloads/owa_attachments/${DATE}"
mkdir -p "${LOG_DIR}"

log_msg() {
  local level="$1" step="$2" msg="$3"
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] [${level}] [${step}] ${msg}" >> "${LOG_FILE}"
}

# 初始化日志
log_msg "INFO" "INIT" "summarise 启动，目标日期: ${DATE}"
log_msg "INFO" "INIT" "附件目录: ${TARGET_DIR}"

# 记录从 email-fetch 日志解析到的附件信息（如果有）
if [ -f "${LOG_DIR}/email-fetch-${EXEC_DATE}.log" ]; then
  DOWNLOADED_COUNT=$(grep -c '\[SUCCESS\] \[DOWNLOAD\] 下载完成:' "${LOG_DIR}/email-fetch-${EXEC_DATE}.log" 2>/dev/null || echo 0)
  FAILED_COUNT=$(grep -c '\[ERROR\] \[DOWNLOAD\] 下载失败:' "${LOG_DIR}/email-fetch-${EXEC_DATE}.log" 2>/dev/null || echo 0)
  log_msg "INFO" "PHASE0" "从 email-fetch 日志解析: 成功下载 ${DOWNLOADED_COUNT} 个, 失败 ${FAILED_COUNT} 个"
fi

# 记录整合结果（整合脚本运行后追加）
# log_msg "SUCCESS" "CONSOLIDATE" "整合完成: 汇总表.xlsx, 728 行, 5 个供应商"

# 记录校验结果
# log_msg "INFO" "VALIDATE" "校验通过: 29 列完整, 无重复行"
# 或
# log_msg "WARN" "VALIDATE" "校验警告: 发现 2 行重复"

log_msg "INFO" "DONE" "执行完成 — 汇总表路径、行数、校验结果"
```

### Phase 4 — 报告

向用户报告：
- ✅ 汇总表格文件路径
- 总行数 / 供应商数量
- 各供应商数据量分布
- 校验结果
- 日志引用

```
📊 汇总完成:
  文件: ~/Downloads/汇总表_20260529.xlsx
  行数: 728 | 供应商: 5 家
  校验: 全部通过 ✅

📋 执行日志: ~/Downloads/owa_attachments/logs/summarise-2026-06-04.log
📁 附件来源: ~/Downloads/owa_attachments/2026-05-29/
```

如果 email-fetch 日志存在，也一并引用：
```
📋 下载日志: ~/Downloads/owa_attachments/logs/email-fetch-2026-06-04.log
```

---

## 输出格式规范

### 29 列（A–AC）

| 列 | 标题 | 列 | 标题 |
|----|------|----|------|
| A | 供应商 | N | 焊线批检 |
| B | 客户代码 | O | 塑封 |
| C | 客户名称 | P | 电镀 |
| D | 客户订单号 | Q | 切筋 |
| E | 封装形式 | R | 委外切筋 |
| F | 产品型号 | S | 测试 |
| G | 芯片型号 | T | 测试编带 |
| H | 晶圆批次 | U | 包装 |
| I | 订单数量 | V | 包装入库 |
| J | 投产日期 | W | 在线合计 |
| K | 粘片 | X | 入库良品 |
| L | 焊线 | Y | 入库不良品 |
| M | 焊线2 | Z | 出库良品 |
| AA | 出库不良品 | AB | 库存 | AC | 状态 |

### 样式

| 元素 | 规范 |
|------|------|
| 表头 | Calibri 11pt 白色加粗，深蓝底 `#2E75B6`，居中，细边框 |
| 数据行 | Calibri 11pt 白色，按供应商分色，居中，细边框 |
| 列宽 | 14（A–AC） |
| 冻结 | A2（冻结表头行） |
| Sheet 名 | `Sheet` |

### 供应商配色逻辑

1. **已知供应商保持原配色**：JC196巨成(紫)、晟矽微尊阳(天蓝)、中芯微(蓝灰)、气派(黄)、内江明泰(绿)
2. **新供应商自动分配**：按数据量排序，从 16 色调色板中依次分配未使用的颜色
3. **供应商别名自动合并**：如"气派"和"气派科技"合并为同一供应商

---

## 与 email-fetch 配合使用

完整流程示例：

1. 用户说："下载今天/昨天的 WIP 邮件附件并整合"
2. 调用 **email-fetch** 下载指定日期邮件附件到 `~/Downloads/owa_attachments/<YYYY-MM-DD>/`，同时写入执行日志
3. summarise 读取 email-fetch 日志解析附件清单（成功/失败），确认文件就绪
4. 运行 **summarise** 脚本整合，传入日期子目录路径
5. summarise 写入自己的执行日志到共享日志目录
6. 向用户报告最终文件路径，并引用两份日志

```
email-fetch → ~/Downloads/owa_attachments/2026-05-29/
            → logs/email-fetch-2026-06-04.log
            ↓
summarise → 解析日志获取附件清单 → consolidate_wip.py
          → logs/summarise-2026-06-04.log
          → ~/Downloads/汇总表.xlsx
```

---

## 新增供应商

如果有新的供应商发来 WIP 表格但脚本尚未支持其格式：

1. 先将该供应商的源文件放入 `~/Downloads/owa_attachments/<YYYY-MM-DD>/`
2. 读取源文件的表头结构和数据行
3. 在 `consolidate_wip.py` 中新增一个解析段，将其列映射到 29 列标准格式
4. 在 `SUPPLIER_ALIASES` 中添加别名映射（如有需要）
5. 运行脚本验证

---

## 日志与 email-fetch 的交互规范

### 读取 email-fetch 日志

summarise 可以解析 email-fetch 日志中的结构化条目，支持：

```bash
# 提取成功下载的完整文件列表（文件名 + 大小）
grep '\[SUCCESS\] \[DOWNLOAD\] 下载完成:' "${LOG_FILE}" | \
  sed -E 's/.*下载完成: (.*)$/\1/'

# 仅提取文件名
grep '\[SUCCESS\] \[DOWNLOAD\] 下载完成:' "${LOG_FILE}" | \
  sed -E 's/.*下载完成: (.*) \(.*/\1/'

# 提取失败的文件及原因
grep '\[ERROR\] \[DOWNLOAD\] 下载失败:' "${LOG_FILE}" | \
  sed -E 's/.*下载失败: (.*) — (.*)$/文件: \1 | 原因: \2/'

# 统计扫描到的邮件数
grep '\[INFO\] \[SCAN\] "扫描完成' "${LOG_FILE}" | \
  sed -E 's/.*匹配到 ([0-9]+) 封.*/\1/'

# 统计下载成功率
TOTAL=$(grep -c '\[DOWNLOAD\]' "${LOG_FILE}" 2>/dev/null || echo 0)
OK=$(grep -c '\[SUCCESS\] \[DOWNLOAD\]' "${LOG_FILE}" 2>/dev/null || echo 0)
echo "成功率: ${OK}/${TOTAL}"
```

### 日志引用

在 Phase 4 报告中，当 email-fetch 日志存在时，始终引用两份日志路径，方便用户追溯完整链路。

### 日志保留

```bash
# 清理 30 天前的日志
find ~/Downloads/owa_attachments/logs/ -name "*.log" -mtime +30 -delete
```

---

## 注意事项

1. **附件格式**：支持 `.xlsx`（openpyxl）和 `.xls`（xlrd）
2. **文件名识别**：脚本通过文件名匹配源文件类型，重命名文件会导致解析失败
3. **表头行定位**：不同源文件的表头行位置不同，脚本中已分别处理
4. **数据量**：单次整合约 700+ 行数据，运行时间约 10-30 秒
5. **日志依赖**：从 email-fetch 日志解析附件时，需确保 email-fetch 日志已完整写入（即 email-fetch 已完成 Step 6）
6. **日志格式稳定性**：email-fetch 日志格式为 `[时间] [级别] [步骤] 消息`，下游解析依赖此结构
