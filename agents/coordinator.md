# 流程参考文档

> 此文件供 main 参考。

## Phase 0: 初始化

### 0.1 创建输出目录

**检查输出目录可写性**（解决权限问题）：
```bash
INPUT_DIR=$(dirname "$(readlink -f '<输入文件>')")
touch "$INPUT_DIR/.write_test" 2>/dev/null && rm "$INPUT_DIR/.write_test" 2>/dev/null
if [ $? -ne 0 ]; then
    OUTPUT_DIR="$HOME/校审_output"
    mkdir -p "$OUTPUT_DIR"
    echo "[WARN] 源目录不可写，输出降级到: $OUTPUT_DIR"
else
    OUTPUT_DIR="$INPUT_DIR"
fi
```

创建结果临时目录：
```bash
mkdir -p /tmp/power-review-results
```

### 0.2 创建团队（可选）

用于 UI 标签页显示，失败不影响后续流程。
```python
TeamCreate(team_name="power-design-review-team", description="输变电工程设计说明校审团队")
```

### 0.3 转换文档

```bash
python3 ~/.agents/tools/doc_to_md.py "<输入文件>" -o "$OUTPUT_DIR/<输出文件.md>"
```

- doc_to_md.py 已内置权限降级逻辑（不可写时自动降级到 `~/校审_output/`）
- 支持 `.pdf` / `.docx` / `.pptx` / `.xlsx` / `.md` 等格式
- `.md` 文件自动透传，其他格式自动转换
- 记录转换后的 MD 文件绝对路径

### 0.4 提取章节索引（多模式匹配）

PDF 转换后的 MD 可能没有标准 `#` 标题，需按优先级尝试多种模式：

```bash
MD_FILE="<转换后的MD文件路径>"

# 模式1: 标准 Markdown 标题（docx 转 MD 时常见）
CHAPTER_INDEX=$(grep -n "^#" "$MD_FILE" | head -100)

# 模式2: 中文章节标题（PDF 转 MD 后常见）
if [ -z "$CHAPTER_INDEX" ]; then
    CHAPTER_INDEX=$(grep -nE "^第[0-9一二三四五六七八九十]+章|^第[0-9]+部分" "$MD_FILE" | head -100)
fi

# 模式3: 纯数字编号标题（如 "1 工程概述"、"2.1 负荷预测"）
if [ -z "$CHAPTER_INDEX" ]; then
    CHAPTER_INDEX=$(grep -nE "^[0-9]+[\.\s]" "$MD_FILE" | head -100)
fi

# 模式4: 都失败则标记为无索引，后续用行号定位
if [ -z "$CHAPTER_INDEX" ]; then
    echo "[WARN] 无法自动提取章节索引，将使用行号定位"
fi
```

将提取到的索引保存，传递给子代理。

### 0.5 评估文档规模

```bash
wc -l "$MD_FILE"
```

| 文档规模 | 行数 | 策略 |
|----------|------|------|
| 小型 | < 2000 行 | 子代理一次性读取全文 |
| 中型 | 2000-8000 行 | 子代理按章节分段读取 |
| 大型 | > 8000 行 | 子代理按章节分段读取，每个专业分批检查 |

### 0.6 启动子代理

**核心变更：子代理通过文件传递结果，不依赖 SendMessage 传递 JSON。**

结果文件约定：
| 角色 | 结果文件路径 |
|------|-------------|
| general-checker | `/tmp/power-review-results/general-checker.json` |
| senior-engineer | `/tmp/power-review-results/senior-engineer.json` |
| 合并结果 | `/tmp/power-review-results/merged.json` |
| reviewer | `/tmp/power-review-results/reviewer.json` |

启动时通过 prompt 传递：
1. **文档路径**：MD 文件的绝对路径
2. **章节索引**：从 0.4 步骤提取的章节结构
3. **文档规模**：从 0.5 步骤判断的规模等级和对应策略
4. **结果文件路径**：明确告知子代理写入哪个文件

**同时启动 general-checker 和 senior-engineer（并行，使用 run_in_background）**：

```python
# 示例：启动 general-checker
Agent(
    name="general-checker",
    description="通用检查-规范术语数据",
    subagent_type="general-purpose",
    team_name="power-design-review-team",
    mode="bypassPermissions",
    prompt="""你是输变电工程设计说明报告的通用检查员。

## 文档信息
- 文档路径: {md_file_path}
- 文档规模: {scale}（{line_count} 行）
- 读取策略: {strategy}

## 章节索引
{chapter_index}

## 结果输出（必须执行）
完成检查后，将 JSON 结果写入文件: /tmp/power-review-results/general-checker.json
然后用 SendMessage 通知 main: "general-checker 检查完成，共 N 个问题"

详细操作请参考 ~/.agents/skills/power-design-review/agents/general-checker.md
"""
)
```

senior-engineer 使用相同格式的 prompt，仅替换角色和结果路径。

---

## Phase 1: 并行检查 + 结果收集

### 1.1 等待子代理完成

等待两个子代理发送完成通知，或等待后台任务完成通知。

### 1.2 从文件读取结果（关键）

**不要从 SendMessage 读取 JSON，从文件读取**：

```bash
# 检查结果文件是否存在且非空
if [ -s /tmp/power-review-results/general-checker.json ]; then
    echo "general-checker 结果已就绪"
    cat /tmp/power-review-results/general-checker.json
else
    echo "[WARN] general-checker 结果文件为空或不存在，需要降级处理"
fi

if [ -s /tmp/power-review-results/senior-engineer.json ]; then
    echo "senior-engineer 结果已就绪"
    cat /tmp/power-review-results/senior-engineer.json
else
    echo "[WARN] senior-engineer 结果文件为空或不存在，需要降级处理"
fi
```

### 1.3 降级策略（关键）

**如果某个子代理的结果文件为空或不存在**，main 自行执行该角色的检查：

1. 读取对应角色定义文件：
   - general-checker → `~/.agents/skills/power-design-review/agents/general-checker.md`
   - senior-engineer → `~/.agents/skills/power-design-review/agents/senior-engineer.md`
2. 按角色职责直接检查文档，生成 JSON 结果
3. 将结果写入对应的结果文件

### 1.4 合并结果

将两个角色的结果合并为一份：
```bash
# 可使用 python 合并 JSON
python3 -c "
import json
gc = json.load(open('/tmp/power-review-results/general-checker.json'))
se = json.load(open('/tmp/power-review-results/senior-engineer.json'))
merged = {'issues': gc.get('issues', []) + se.get('issues', [])}
json.dump(merged, open('/tmp/power-review-results/merged.json', 'w'), ensure_ascii=False, indent=2)
print(f'合并完成: {len(merged[\"issues\"])} 个问题')
"
```

---

## Phase 2: 审查反馈

### 2.1 启动 reviewer

通过 prompt 传递文件路径，reviewer 从文件读取合并结果：

```python
Agent(
    name="reviewer",
    description="审查输出并修正位置格式",
    subagent_type="general-purpose",
    team_name="power-design-review-team",
    mode="bypassPermissions",
    prompt="""你是输变电工程校审团队的审查员。

## 输入文件
- 合并结果: /tmp/power-review-results/merged.json
- 文档路径: {md_file_path}
- 章节索引:
{chapter_index}

## 结果输出（必须执行）
完成审查后，将修正后的 JSON 结果写入: /tmp/power-review-results/reviewer.json
然后用 SendMessage 通知 main: "reviewer 审查完成，修正 M 个位置，共 N 个问题"

详细操作请参考 ~/.agents/skills/power-design-review/agents/reviewer.md
"""
)
```

### 2.2 读取 reviewer 结果

```bash
cat /tmp/power-review-results/reviewer.json
```

### 2.3 reviewer 降级

如果 reviewer 结果文件为空或不存在：
1. main 自行执行行号→章节编号的修正
2. 使用章节索引手动映射行号到章节
3. 检查分类名称是否在规定范围内

---

## Phase 3: 汇总输出（全自动，无需用户干预）

1. 使用 reviewer 返回的修正后问题清单（或降级后 main 自行修正的结果）
2. 合并问题，去重
3. 生成 MD 文档
4. **自动生成 Word 文档（不要询问用户）**：
   ```bash
   python3 ~/.agents/skills/power-design-review/scripts/md_to_word.py "<MD文件路径>" --force
   ```

**关键指令**：
- **不要询问用户是否生成 Word 文档** — 这是流程的必需步骤
- 使用**绝对路径**调用脚本
- 使用 **--force** 参数确保转换成功
- 如果转换失败，在输出中说明原因，但**不要中断流程**

---

## 清理

```bash
# 清理临时结果文件
rm -f /tmp/power-review-results/*.json

# 清理团队（如有活跃代理无法正常关闭，可强制清理）
rm -rf ~/.claude/teams/power-design-review-team ~/.claude/tasks/power-design-review-team
```

---

## 输出验证清单

- [ ] 文档已成功转换为 MD
- [ ] 输出目录可写（或已降级）
- [ ] 章节索引已提取（无论哪种模式）
- [ ] general-checker 结果已获取（文件读取或降级）
- [ ] senior-engineer 结果已获取（文件读取或降级）
- [ ] reviewer 结果已获取（文件读取或降级）
- [ ] 问题位置全部为章节编号格式
- [ ] MD文档已生成
- [ ] Word文档已自动生成
- [ ] 临时文件已清理

## 异常处理

| 异常场景 | 处理方式 |
|----------|----------|
| 输出目录不可写 | doc_to_md.py 自动降级到 `~/校审_output/` |
| 子代理结果文件为空 | main 自行执行该角色检查（降级模式） |
| 子代理无法关闭 | `rm -rf ~/.claude/teams/...` 强制清理 |
| 章节索引为空 | 使用行号定位，reviewer 或 main 手动修正 |
| PDF 转换内容为空 | 提示用户可能是扫描件，需 OCR |
| Word 生成失败 | 输出 MD 文件，说明 Word 生成失败原因 |
