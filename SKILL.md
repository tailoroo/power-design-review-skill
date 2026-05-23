---
name: power-design-review
description: 输变电工程设计说明校审。支持初步设计、施工图设计等阶段。检查内容矛盾、数据不一致、术语错误、规范合规性。输出5列表格问题清单及Word文档。
---

# 输变电工程设计说明校审

## 一、何时使用

校审输变电工程设计说明报告时触发此skill：
- 收到初步设计说明报告
- 收到施工图设计说明报告
- 需要检查内容矛盾、数据不一致、术语错误、规范合规性

**支持多种输入格式**：
- `.docx` — 自动转换为 Markdown 后校审（推荐）
- `.pdf` — 文字型 PDF 自动转换，扫描件需 OCR
- `.pptx` / `.xlsx` 等 — 通过 markitdown 自动转换
- `.md` — 直接校审

输入转换使用通用工具 `~/.agents/tools/doc_to_md.py`

## 二、架构（3子代理 + 文件传递）

```
                    ┌─────────────────────────┐
                    │         main            │
                    │       (总协调者)         │
                    └──────────┬──────────────┘
                               │
        ┌──────────────────────┼──────────────────────┐
        ▼                      ▼                      ▼
┌───────────────┐     ┌───────────────┐     ┌───────────────┐
│general-checker│     │senior-engineer│     │   reviewer    │
│   通用检查员   │     │ 资深专业工程师 │     │    审查员     │
└───────┬───────┘     └───────┬───────┘     └───────────────┘
        │                     │                     ▲
        │  /tmp/.../gc.json   │  /tmp/.../se.json   │
        └─────────────────────┴──────────┬──────────┘
                                           │
                            /tmp/.../merged.json
                            /tmp/.../reviewer.json
```

**结果传递机制**：子代理将 JSON 结果写入 `/tmp/power-review-results/` 目录，main 从文件读取。SendMessage 仅用于发送简短完成通知（不含 JSON）。

### 角色职责

| 角色 | 职责 |
|------|------|
| **main** | 读取文档、分配任务、汇总结果、生成文档、降级执行 |
| **general-checker** | 规范版本、术语、数据一致性、站名统一 |
| **senior-engineer** | 系统专业、电气一次/二次、土建、概算 |
| **reviewer** | 审核输出、**修正问题位置格式**、发现遗漏 |

### 职责边界说明

| 检查类型 | 负责角色 | 说明 |
|----------|----------|------|
| 规范版本核实 | general-checker | 检查引用标准是否过期 |
| 术语/错别字 | general-checker | 对照 common-errors.json 检查 |
| 深度技术审查 | senior-engineer | 系统方案、设备选型等专业内容 |
| 输出审核 | reviewer | 审核其他角色输出质量 |
| 位置格式修正 | reviewer | 将行号格式改为章节编号 |

### 各角色禁止事项

| 角色 | 不做什么 |
|------|----------|
| **general-checker** | ① 不做深度专业技术审查（由 senior-engineer 负责）<br>② 不审查其他角色输出（由 reviewer 负责）<br>③ 不记录 Word转MD 格式问题（如 LaTeX 公式残留） |
| **senior-engineer** | ① 不检查规范版本（由 general-checker 负责）<br>② 不检查错别字（由 general-checker 负责）<br>③ 不审查其他角色输出（由 reviewer 负责） |
| **reviewer** | ① 不直接执行内容检查<br>② 不保留行号格式（必须修正为章节编号） |

## 三、执行流程

```
Phase 0: 初始化
├── 0.1 检查输出目录权限，不可写则降级到 ~/校审_output/
├── 0.2 创建团队（可选，用于 UI 标签页）
├── 0.3 转换文档（doc_to_md.py，内置权限降级）
├── 0.4 提取章节索引（多模式匹配：# 标题 → 中文章节 → 数字编号）
├── 0.5 评估文档规模（wc -l）
├── 创建 /tmp/power-review-results/ 临时目录
└── 0.6 启动 2 个检查子代理（并行），指定结果文件路径

Phase 1: 并行检查 + 结果收集
├── general-checker 和 senior-engineer 并行执行
├── 子代理将结果写入 /tmp/power-review-results/*.json
├── main 从文件读取结果（不依赖 SendMessage 传 JSON）
├── 如果结果文件为空 → main 降级自行执行该角色检查
└── 合并两个角色的结果为 merged.json

Phase 2: 审查反馈
├── 启动 reviewer，从文件读取 merged.json
├── reviewer 审核检查结果，修正行号→章节编号
├── reviewer 将结果写入 /tmp/power-review-results/reviewer.json
├── main 从文件读取 reviewer 结果
└── 如果 reviewer 结果为空 → main 自行修正

Phase 3: 汇总输出（全自动，无需用户干预）
├── 使用 reviewer 修正后的问题清单
├── 合并问题，去重
├── 生成 MD 文档
├── 自动生成 Word 文档（--force，不询问用户）
└── 清理临时文件和团队
```

### 长文档支持

| 文档规模 | 行数 | 读取策略 |
|----------|------|----------|
| 小型 | < 2000 行 | 一次性读取 |
| 中型 | 2000-8000 行 | 按章节分段读取（每段 300-800 行） |
| 大型 | > 8000 行 | 按专业对应章节分批读取检查 |

### 降级策略

| 故障场景 | 降级方案 |
|----------|----------|
| 子代理结果文件为空 | main 自行执行该角色检查 |
| 子代理无法关闭 | 强制 `rm -rf ~/.claude/teams/...` |
| 输出目录不可写 | 自动降级到 `~/校审_output/` |
| 章节索引提取失败 | 使用行号定位 |
| PDF 转换内容为空 | 提示用户可能是扫描件 |

## 四、问题位置格式

- 正确：`第3章第2.1节`、`第4章表4-1`
- 错误：`第3035行`（reviewer会修正）

## 五、输出格式

5列表格：序号 | 分类 | 问题详细内容 | 问题位置 | 修改建议

分类：系统、电气一次、电气二次、土建、水暖、概算、综合、其他

## 六、文件参考

| 文件 | 说明 |
|------|------|
| `agents/coordinator.md` | 流程参考文档（含降级策略） |
| `agents/general-checker.md` | 通用检查员职责 |
| `agents/senior-engineer.md` | 资深工程师职责 |
| `agents/reviewer.md` | 审查员职责 |
| `references/common-errors.json` | 常见错误配置 |
| `references/review-checklist.md` | 检查清单 |
| `scripts/md_to_word.py` | MD转Word脚本 |
| `scripts/docx_to_md.py` | Word转MD脚本（native 备选） |
| `~/.agents/tools/doc_to_md.py` | 通用文档转MD工具（所有 skill 共享） |

## 七、输出验证清单

- [ ] 输出目录可写（或已降级）
- [ ] 文档已成功转换为 MD
- [ ] 章节索引已提取（无论哪种模式）
- [ ] general-checker 结果已获取（文件读取或降级）
- [ ] senior-engineer 结果已获取（文件读取或降级）
- [ ] reviewer 结果已获取（文件读取或降级）
- [ ] 问题位置全部为章节编号格式
- [ ] MD文档已生成
- [ ] Word文档已自动生成（无需询问用户）
- [ ] 临时文件已清理

## 八、子代理调用规范

启动子代理时**必须使用 `name` 参数**和 **`mode="bypassPermissions"`**：

```python
# 启动 general-checker
Agent(
    name="general-checker",
    description="通用检查-规范术语数据",
    subagent_type="general-purpose",
    team_name="power-design-review-team",
    mode="bypassPermissions",
    prompt="""...（含文档路径、章节索引、规模、结果文件路径）..."""
)

# 启动 senior-engineer
Agent(
    name="senior-engineer",
    description="专业深度检查",
    subagent_type="general-purpose",
    team_name="power-design-review-team",
    mode="bypassPermissions",
    prompt="""...（含文档路径、章节索引、规模、结果文件路径）..."""
)

# Phase 2 启动 reviewer
Agent(
    name="reviewer",
    description="审查输出并修正位置格式",
    subagent_type="general-purpose",
    team_name="power-design-review-team",
    mode="bypassPermissions",
    prompt="""...（含合并结果文件路径、文档路径、章节索引）..."""
)
```

**关键**：
- 子代理必须将 JSON 结果写入 `/tmp/power-review-results/` 目录
- SendMessage 仅用于简短通知（如"检查完成，共 N 个问题"），不传 JSON
- 如果子代理结果文件为空，main 降级自行执行

## 九、清理

```bash
# 清理临时结果文件
rm -f /tmp/power-review-results/*.json

# 清理团队（代理无法关闭时强制清理）
rm -rf ~/.claude/teams/power-design-review-team ~/.claude/tasks/power-design-review-team
```
