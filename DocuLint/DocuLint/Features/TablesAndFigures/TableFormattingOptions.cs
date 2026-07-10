namespace DocuLint
{
    internal sealed class TableFormattingOptions
    {
        public string HeaderFontName { get; set; }
        public float HeaderFontSizePoints { get; set; }
        public string BodyFontName { get; set; }
        public float BodyFontSizePoints { get; set; }
        public float TableWidthCentimeters { get; set; }
        public float OuterBorderWidthPoints { get; set; }
        public float InnerBorderWidthPoints { get; set; }

        public static TableFormattingOptions CreateDefault()
        {
            return new TableFormattingOptions
            {
                HeaderFontName = "黑体",
                HeaderFontSizePoints = 12f,
                BodyFontName = "宋体",
                BodyFontSizePoints = 10.5f,
                TableWidthCentimeters = 17.4f,
                OuterBorderWidthPoints = 1.5f,
                InnerBorderWidthPoints = 0.5f
            };
        }

        public TableFormattingOptions Clone()
        {
            return new TableFormattingOptions
            {
                HeaderFontName = HeaderFontName,
                HeaderFontSizePoints = HeaderFontSizePoints,
                BodyFontName = BodyFontName,
                BodyFontSizePoints = BodyFontSizePoints,
                TableWidthCentimeters = TableWidthCentimeters,
                OuterBorderWidthPoints = OuterBorderWidthPoints,
                InnerBorderWidthPoints = InnerBorderWidthPoints
            };
        }
    }
}