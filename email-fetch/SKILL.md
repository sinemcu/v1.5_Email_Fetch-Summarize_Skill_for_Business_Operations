---
name: email-fetch
description: >
  Download attachments from emails on a specific date via OWA (Outlook Web App) at
  https://mail.sinomcu.com/owa. Use when the user asks to download email attachments
  from OWA, fetch attachments from a given date, or save WIP reports/attachments from
  the company email inbox. Always confirm the target date with the user before starting.
---

# Email Fetch — OWA Attachment Downloader

## Workflow

```
1. Confirm target date with user
2. Navigate to OWA and login if needed (browser action navigate + snapshot + act)
3. Scan mail list, collect all items matching the target date
4. For each matching email → open → extract attachment URLs → download via JS fetch()
5. Copy downloaded files from /tmp/openclaw/downloads/ → ~/Downloads/owa_attachments/<YYYY-MM-DD>/
6. Write structured execution log to ~/Downloads/owa_attachments/logs/
7. Report final file list
```

## Environment

| Item | Value |
|------|-------|
| URL | `https://mail.sinomcu.com/owa` |
| Browser download dir | `/tmp/openclaw/downloads/` |
| Target base dir | `~/Downloads/owa_attachments/` |
| Date subdirectory | `~/Downloads/owa_attachments/<YYYY-MM-DD>/` |
| Log directory | `~/Downloads/owa_attachments/logs/` |
| Log file | `~/Downloads/owa_attachments/logs/email-fetch-YYYY-MM-DD.log`（按执行日期命名） |
| Email Account | 由用户在运行时提供（见下方 Credentials 说明） |
| Password | 由用户在运行时提供（见下方 Credentials 说明） |

## Step 1 — Confirm Date

Ask the user: "确认一下，你要下载 **YYYY-MM-DD** 的邮件附件？"

Proceed only after user confirms.

Initialize the log directory and helper:

```bash
EXEC_DATE="$(date '+%Y-%m-%d')"  # 执行日期（北京时间）
LOG_DIR="$HOME/Downloads/owa_attachments/logs"
LOG_FILE="${LOG_DIR}/email-fetch-${EXEC_DATE}.log"
mkdir -p "${LOG_DIR}"

log_msg() {
  local level="$1" step="$2" msg="$3"
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] [${level}] [${step}] ${msg}" >> "${LOG_FILE}"
}

# Initial log entries
log_msg "INFO" "INIT" "email-fetch 启动，目标邮件日期: ${DATE}"
log_msg "INFO" "INIT" "执行日期: ${EXEC_DATE}，日志文件: ${LOG_FILE}"
```

## Step 2 — Navigate & Login

### 2b — 读取 credentials

在执行导航前，先按优先级尝试获取账号密码：

```bash
# Step 1: 从 password.json 读取（SKILL.md 同级目录）
SKILL_DIR="$(cd "$(dirname "$(readlink -f ~/.openclaw/workspace/skills/email-fetch/SKILL.md)")" && pwd)"
PASSWORD_JSON="${SKILL_DIR}/password.json"
CRED_ACCOUNT=""
CRED_PASSWORD=""

if [ -f "${PASSWORD_JSON}" ]; then
  CRED_ACCOUNT=$(python3 -c "import json; d=json.load(open('${PASSWORD_JSON}')); print(d.get('account',''))" 2>/dev/null)
  CRED_PASSWORD=$(python3 -c "import json; d=json.load(open('${PASSWORD_JSON}')); print(d.get('password',''))" 2>/dev/null)
fi

# Step 2: 如果 password.json 为空，回退到环境变量
if [ -z "${CRED_ACCOUNT}" ]; then
  CRED_ACCOUNT="${OWA_ACCOUNT:-}"
  CRED_PASSWORD="${OWA_PASSWORD:-}"
fi

if [ -n "${CRED_ACCOUNT}" ] && [ -n "${CRED_PASSWORD}" ]; then
  log_msg "INFO" "CREDS" "已从 password.json 或环境变量读取账号: ${CRED_ACCOUNT}"
else
  log_msg "INFO" "CREDS" "未找到凭据，需要用户在对话中提供"
  # 在对话中向用户请求账号密码
fi
```

使用 `${CRED_ACCOUNT}` 和 `${CRED_PASSWORD}` 填入 OWA 登录表单。如果两者都为空，则在对话中向用户请求。

1. Navigate to OWA:
```
browser → action: navigate → url: https://mail.sinomcu.com/owa
```

2. Check if already logged in:
```
browser → action: snapshot
```
- If the page shows the mail inbox (邮件列表 visible), skip login and proceed to Step 3.
- If the page shows the OWA login screen, perform login below.

After navigation, log:
```bash
log_msg "INFO" "NAVIGATE" "已导航至 https://mail.sinomcu.com/owa"
log_msg "INFO" "NAVIGATE" "页面加载完成，检查登录状态"
```

### 2a — OWA Login

**Credentials source (in order of preference):**
1. Read from `password.json` in the same directory as this SKILL.md — if both `account` and `password` are non-empty strings, use them
2. User provides account & password directly in the conversation
3. Read from environment variables `OWA_ACCOUNT` and `OWA_PASSWORD` if set
4. Ask the user to provide them interactively

`password.json` 格式：
```json
{
  "account": "your-email@sinomcu.com",
  "password": "your-password"
}
```

⚠️ **安全注意**：`password.json` 为明文存储，仅限本地信任环境使用。如无需自动登录，可留空或删除该文件。

Login flow:
```
browser → action: snapshot  (find email input and password input refs)
browser → action: act → ref=<email_input_ref> → kind=fill → text=<account>
browser → action: act → ref=<password_input_ref> → kind=fill → text=<password>
browser → action: act → ref=<sign_in_button_ref> → kind=click
```
Wait 5 seconds for login to complete, then take another snapshot to confirm inbox is loaded.

If login succeeds:
```bash
log_msg "SUCCESS" "LOGIN" "登录成功，邮箱已就绪"
```

If login fails:
```bash
log_msg "ERROR" "LOGIN" "登录失败，请检查账号密码"
```

Report the error and ask the user to check credentials.

## Step 3 — Find Matching Emails

Run this JS to scan the mail list and find items whose text contains the target date (supports formats like `5/17`, `2026-05-17`, `周一 5/17`, etc.):

```javascript
// Evaluate in browser, pass the date string to search for
(targetDate) => {
  const mailList = document.querySelector('[aria-label="邮件列表"]');
  if (!mailList) return 'mail list not found';
  const items = Array.from(mailList.querySelectorAll('div[role="button"]'));

  const matched = [];
  for (let i = 0; i < items.length; i++) {
    const text = items[i].textContent;
    if (text.includes(targetDate)) {
      matched.push({
        index: i,
        text: text.replace(/\s+/g, ' ').trim().substring(0, 120)
      });
    }
  }
  return JSON.stringify(matched);
}
```

If no matches found, try alternate date formats (e.g., `5/17` vs `2026-05-17` vs `周日 5/17`).

After scanning, log:
```bash
log_msg "INFO" "SCAN" "扫描完成，匹配到 N 封邮件"
# 逐封记录
log_msg "INFO" "SCAN" "[0] 邮件主题摘要..."
log_msg "INFO" "SCAN" "[1] 邮件主题摘要..."
```

If no matches:
```bash
log_msg "WARN" "SCAN" "未找到目标日期的邮件，已尝试多种日期格式"
```

## Step 4 — Download Attachments from Each Email

### 4a — Open email and verify attachment

```javascript
// Open email at given index
(items) => {
  const item = items[targetIndex];
  item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, view: window }));
  item.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, view: window }));
  item.dispatchEvent(new MouseEvent('click', { bubbles: true, view: window }));
  return 'clicked';
}
```

Wait 3 seconds, then check reading pane:

```javascript
() => {
  const readingPane = document.querySelector('[aria-label="阅读窗格"]');
  if (!readingPane) return { heading: 'no reading pane' };
  const heading = readingPane.querySelector('h1, h2, [role="heading"]');
  const links = Array.from(readingPane.querySelectorAll('a'));
  const attLinks = links.filter(l => l.href && l.href.includes('GetFileAttachment') && l.textContent.trim());
  return { heading: heading ? heading.textContent : 'none', attachments: attLinks.map(l => l.textContent.trim()) };
}
```

### 4b — Download all attachments from current email

```javascript
async () => {
  const links = Array.from(document.querySelectorAll('a'));
  const attLinks = links.filter(l => l.href && l.href.includes('GetFileAttachment') && l.textContent.trim());

  const results = [];
  for (const link of attLinks) {
    try {
      const resp = await fetch(link.href);
      const blob = await resp.blob();

      // Try to get real filename from URL params or Content-Disposition header
      let filename = null;
      const urlParams = new URLSearchParams(link.href.split('?')[1]);
      filename = urlParams.get('name') || urlParams.get('filename');

      // Fallback: visible text (may be truncated)
      if (!filename) filename = link.textContent.trim().replace(/[~]/g, '_');

      // Sanitize
      filename = filename.replace(/[~\\/\\?%*:|"<>]/g, '_');

      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      results.push({ filename, size: blob.size });
      await new Promise(r => setTimeout(r, 500));
    } catch(e) {
      results.push({ error: e.message, link: link.textContent.trim() });
    }
  }
  return JSON.stringify(results);
}
```

### 4c — Reusable single-script (download all attachments for one email by index)

For convenience, this combined script opens an email and downloads all its attachments in one call:

```javascript
async (emailIndex) => {
  const items = Array.from(document.querySelectorAll('div[role="button"]'));
  const item = items[emailIndex];
  if (!item) return { error: `Email at index ${emailIndex} not found` };

  item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, view: window }));
  item.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, view: window }));
  item.dispatchEvent(new MouseEvent('click', { bubbles: true, view: window }));

  await new Promise(r => setTimeout(r, 3000));

  const links = Array.from(document.querySelectorAll('a'));
  const attLinks = links.filter(l => l.href && l.href.includes('GetFileAttachment') && l.textContent.trim());

  const results = [];
  for (const link of attLinks) {
    try {
      const resp = await fetch(link.href);
      const blob = await resp.blob();

      // Try to get real filename from URL params
      let filename = null;
      const urlParams = new URLSearchParams(link.href.split('?')[1]);
      filename = urlParams.get('name') || urlParams.get('filename');

      if (!filename) filename = link.textContent.trim().replace(/[~]/g, '_');
      filename = filename.replace(/[~\\/\\?%*:|"<>]/g, '_');

      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      results.push({ filename, size: blob.size });
      await new Promise(r => setTimeout(r, 500));
    } catch(e) {
      results.push({ error: e.message, link: link.textContent.trim() });
    }
  }
  return { emailIndex, attachments: results };
}
```

After each email, log results:
```bash
log_msg "INFO" "DOWNLOAD" "打开邮件 [0]: \"WIP报告 - 供应商A\""
# 对每个附件:
log_msg "SUCCESS" "DOWNLOAD" "下载完成: WIP_20260529.xlsx (245KB)"
log_msg "ERROR" "DOWNLOAD" "下载失败: report_v2.pdf — fetch timeout"
```

## Step 5 — Copy Files to Date Subdirectory

After all downloads are triggered:

```bash
# Date format: YYYY-MM-DD (e.g. 2026-05-29)
DATE="2026-05-29"  # use the confirmed target date
TARGET_DIR="~/Downloads/owa_attachments/${DATE}"

mkdir -p "${TARGET_DIR}"

# Find and copy the latest downloaded files (created within last 5 minutes)
find /tmp/openclaw/downloads/ -type f -mmin -5 -exec cp {} "${TARGET_DIR}/" \;

# Strip UUID prefixes from filenames
# Pattern: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX-filename.ext
# or partial: XXXXXXXXXXXX-filename.ext (remaining segment)
cd "${TARGET_DIR}"
for f in [0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f]-*; do
  newname="${f#????????????-}"
  if [ "$f" != "$newname" ]; then
    mv "$f" "$newname"
  fi
done

# Count and log
FILE_COUNT=$(ls -1 "${TARGET_DIR}" | wc -l | tr -d ' ')
log_msg "INFO" "COPY" "已拷贝 ${FILE_COUNT} 个文件至 ${TARGET_DIR}"
```

## Step 6 — Finalize Log

Write execution summary to the log:

```bash
# 统计成功/失败数量（从日志中计数）
SUCCESS_COUNT=$(grep -c '\[SUCCESS\] \[DOWNLOAD\]' "${LOG_FILE}" 2>/dev/null || echo 0)
ERROR_COUNT=$(grep -c '\[ERROR\] \[DOWNLOAD\]' "${LOG_FILE}" 2>/dev/null || echo 0)
EMAIL_COUNT=$(grep -c '\[INFO\] \[DOWNLOAD\] "打开邮件' "${LOG_FILE}" 2>/dev/null || echo 0)

log_msg "INFO" "DONE" "执行完成 — 扫描邮件: ${EMAIL_COUNT} 封, 成功下载: ${SUCCESS_COUNT} 个, 失败: ${ERROR_COUNT} 个"
log_msg "INFO" "DONE" "附件目录: ${TARGET_DIR}"
```

## Step 7 — Report

Present a summary table:

```
| 来源邮件 | 附件文件 | 大小 |
|---------|---------|------|
| ... | ... | ... |
```

Also report paths:
> 📁 附件保存位置：`~/Downloads/owa_attachments/<YYYY-MM-DD>/`
> 📋 执行日志：`~/Downloads/owa_attachments/logs/email-fetch-YYYY-MM-DD.log`

## Notes

1. **Session**: Browser must be logged into OWA. Step 2 includes automatic login detection and execution.
2. **Credentials**: Account and password are provided at runtime — never hardcoded in this skill file. Prefer passing them in the conversation or via environment variables.
3. **OWA dynamic DOM**: Always use `evaluate` with JS queries, not static element refs.
4. **Attachment URL expiry**: `X-OWA-CANARY` tokens expire — download immediately after extracting URL.
5. **Filename cleanup**: OWA filenames often contain `~` — always replace with `_`.
6. **Wait 3s** after clicking each email to let the reading pane load.
7. **Log format**: `[YYYY-MM-DD HH:MM:SS] [LEVEL] [STEP] 消息` — 下游技能（如 summarise）可解析此格式读取下载记录。
8. **Log retention**: 建议定期清理 30 天前的日志文件：
   ```bash
   find ~/Downloads/owa_attachments/logs/ -name "email-fetch-*.log" -mtime +30 -delete
   ```
