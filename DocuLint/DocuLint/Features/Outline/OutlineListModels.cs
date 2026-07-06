using System.Collections.Generic;
using System.Linq;

namespace DocuLint
{
    // 编号格式枚举：定义两种大纲编号样式
    public enum OutlineNumberPattern
    {
        Decimal,        // 普通数字格式：1、1.1、1.1.1
        Parenthesized   // 带括号格式：(1)、(2)
    }

    // 大纲重建选项：存储批量重编编号的所有设置
    public class OutlineListRebuildOptions
    {
        // 用户选择要重新编号的级别（如1级、2级、3级标题）
        public HashSet<int> SelectedLevels { get; set; } = new HashSet<int>();

        // 编号样式（普通数字 / 括号数字）
        public OutlineNumberPattern NumberPattern { get; set; }

        // 对齐方式
        public int Alignment { get; set; }

        // 编号后面跟随的字符（空格、制表符等）
        public int TrailingCharacter { get; set; }

        // 是否清除文档里手动输入的编号
        public bool ClearManualNumbering { get; set; } = true;

        // 获取选中的最大级别（如选了1、2、3级，返回3）
        public int MaxSelectedLevel => SelectedLevels.Count == 0 ? 0 : SelectedLevels.Max();
    }

    // 大纲重建结果：执行完批量编号后，返回统计信息
    public class OutlineRebuildResult
    {
        // 已选择处理的级别列表
        public IReadOnlyList<int> SelectedLevels { get; set; }

        // 扫描范围（整篇文档 / 所选内容）
        public string ScanScope { get; set; }

        // 目标处理的段落总数
        public int TargetParagraphCount { get; set; }

        // 清理的列表数量
        public int ClearedListCount { get; set; }

        // 清理的手动编号数量
        public int ClearedManualNumberCount { get; set; }

        // 成功应用样式的段落数
        public int AppliedParagraphCount { get; set; }

        // 扫描耗时（毫秒）
        public long ScanMilliseconds { get; set; }

        // 清理耗时（毫秒）
        public long CleanupMilliseconds { get; set; }

        // 应用耗时（毫秒）
        public long ApplyMilliseconds { get; set; }

        // 总耗时（毫秒）
        public long DurationMilliseconds { get; set; }

        // 各级别关联的样式名称
        public Dictionary<int, string> LinkedStyles { get; set; }
    }
}