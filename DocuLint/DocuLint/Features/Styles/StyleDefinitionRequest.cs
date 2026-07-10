using System.Collections.Generic;

namespace DocuLint
{
    internal sealed class StyleDefinitionRequest
    {
        public int Level { get; set; }

        public int OutlineLevel { get; set; }

        public bool ShouldCreate { get; set; }

        public string StyleName { get; set; }

        public string FontName { get; set; }

        public float FontSize { get; set; }

        public string ListFontName { get; set; }

        public float ListFontSize { get; set; }

        public bool Bold { get; set; }

        public int Alignment { get; set; }

        public float LineSpacing { get; set; }

        public static List<StyleDefinitionRequest> CreateDefaultSet()
        {
            List<StyleDefinitionRequest> items = new List<StyleDefinitionRequest>();
            for (int level = 1; level <= 9; level++)
            {
                items.Add(new StyleDefinitionRequest
                {
                    Level = level,
                    OutlineLevel = level,
                    ShouldCreate = false,
                    StyleName = GetDefaultStyleName(level),
                    FontName = level == 1 ? "黑体" : "宋体",
                    FontSize = 12f,
                    ListFontName = level == 1 ? "黑体" : "宋体",
                    ListFontSize = 12f,
                    Bold = false,
                    Alignment = 0,
                    LineSpacing = 20f
                });
            }

            items.Add(new StyleDefinitionRequest
            {
                Level = 10,
                OutlineLevel = 10,
                ShouldCreate = false,
                StyleName = "正文",
                FontName = "宋体",
                FontSize = 12f,
                ListFontName = "宋体",
                ListFontSize = 12f,
                Bold = false,
                Alignment = 0,
                LineSpacing = 20f
            });
            return items;
        }

        public static string GetDefaultStyleName(int level)
        {
            switch (level)
            {
                case 1:
                    return "通用1级标题";
                case 2:
                    return "通用2级标题";
                case 3:
                    return "通用3级标题";
                case 4:
                    return "通用4级标题";
                case 5:
                    return "通用5级标题";
                case 6:
                    return "通用6级标题";
                case 7:
                    return "通用7级标题";
                case 8:
                    return "通用8级标题";
                case 9:
                    return "通用9级标题";
                default:
                    return "正文";
            }
        }
    }
}
