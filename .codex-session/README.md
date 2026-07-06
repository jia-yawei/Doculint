# 本地会话管理

这个目录用于保存 Codex / AI 协作的本地会话记录。

约定如下：

- 每次开始新一轮工作前，先读取 `session-brief.local.md`。
- 每次结束一轮工作后，追加一份 `history\yyyyMMdd-HHmmss.md` 记录。
- `session-brief.local.md` 只保留“当前最重要的上下文摘要”。
- `history\` 下保存完整的分轮记录，但默认不纳入 Git。

推荐流程：

1. 启动时运行 `scripts/session/start-session.ps1`
2. 结束时运行 `scripts/session/save-session.ps1`
3. 如果阶段目标变化，再手动更新 `session-brief.local.md`
