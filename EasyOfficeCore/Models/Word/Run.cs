using System.Collections.Generic;

namespace EasyOffice.Models.Word
{
    public class Run
    {
        public Run()
        {

        }
        public string Text { get; set; } = "";
        public string Color { get; set; } = "black";
        public int FontSize { get; set; } = 12;
        public string FontFamily { get; set; } = "等线 (中文正文)";
        public bool IsBold { get; set; } = false;
        public UnderlineType UnderlineType { get; set; } = UnderlineType.None;
        /// <summary>
        /// 斜体
        /// </summary>
        public bool IsItalic { get; set; } = false;
        public bool HasLineBreak { get; set; } = false;
        public List<Picture> Pictures { get; set; }
    }

}
