using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SuppliesReport.Utility
{
    public class CustomFormat
    {
        //string width, 
        //JustificationValues justificationValue, TableVerticalAlignmentValues tableVerticalAlignmentValue,
        //string lineSpacing, LineSpacingRuleValues lineSpacingRule,
        //string fontSizeString, string fontString

        private float _cellWidth;

        public float CellWidth
        {
            get { return _cellWidth; }
            set { _cellWidth = value; }
        }

        public string CellWidthString
        {
            get
            {
                if (_cellWidth == -1)
                    return string.Empty;
                else
                    return ((int)(_cellWidth * 300)).ToString();
            }
        }

        private JustificationValues _justificationValue;

        public JustificationValues JustificationValue
        {
            get { return _justificationValue; }
            set { _justificationValue = value; }
        }

        private TableVerticalAlignmentValues _tableVerticalAlignmentValue;

        public TableVerticalAlignmentValues TableVerticalAlignmentValue
        {
            get { return _tableVerticalAlignmentValue; }
            set { _tableVerticalAlignmentValue = value; }
        }

        private string _lineSpacing;

        public string LineSpacing
        {
            get { return _lineSpacing; }
            set { _lineSpacing = value; }
        }

        private LineSpacingRuleValues _lineSpacingRule;

        public LineSpacingRuleValues LineSpacingRule
        {
            get { return _lineSpacingRule; }
            set { _lineSpacingRule = value; }
        }

        private string _lineSpacingBefore;

        public string LineSpacingBefore
        {
            get { return _lineSpacingBefore; }
            set { _lineSpacingBefore = value; }
        }

        private string _lineSpacingAfter;

        public string LineSpacingAfter
        {
            get { return _lineSpacingAfter; }
            set { _lineSpacingAfter = value; }
        }

        private float _fontSize;

        public float FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; }
        }

        public string FontSizeString
        {
            get { return ((int)_fontSize * 2).ToString(); }
        }

        private string _font;

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }


        public CustomFormat(float cellWidth, JustificationValues justification, TableVerticalAlignmentValues tableVerticalAlignment,
            string lineSpacing, LineSpacingRuleValues lineSpacingRule, string lineSpacingBefore, string lineSpacingAfter, 
            string font, float fontSize)
        {
            CellWidth = cellWidth;
            JustificationValue = justification;
            TableVerticalAlignmentValue = tableVerticalAlignment;
            LineSpacing = lineSpacing;
            LineSpacingRule = lineSpacingRule;
            LineSpacingBefore = lineSpacingBefore;
            LineSpacingAfter = lineSpacingAfter;
            FontSize = fontSize;
            Font = font;
        }

        /// <summary>
        /// Constructor with by default:
        ///     TableVerticalAlignmentValue = TableVerticalAlignmentValues.Center;
        ///     LineSpacing = "240";
        ///     LineSpacingRule = LineSpacingRuleValues.Auto;
        ///     LineSpacingBefore = "0";
        ///     LineSpacingAfter = "0";
        ///     Font = "Times New Roman";
        /// </summary>
        /// <param name="cellWidth"></param>
        /// <param name="justification"></param>
        /// <param name="fontSize"></param>
        public CustomFormat(float cellWidth, JustificationValues justification, float fontSize)
        {
            CellWidth = cellWidth;
            JustificationValue = justification;
            TableVerticalAlignmentValue = TableVerticalAlignmentValues.Center;
            LineSpacing = "240";
            LineSpacingRule = LineSpacingRuleValues.Auto;
            LineSpacingBefore = "0";
            LineSpacingAfter = "0";
            FontSize = fontSize;
            Font = "Times New Roman";
        }

        /// <summary>
        /// Constructor with by default:
        ///     CellWidth = -1;
        ///     TableVerticalAlignmentValue = TableVerticalAlignmentValues.Center;
        ///     LineSpacing = "240";
        ///     LineSpacingRule = LineSpacingRuleValues.Auto;
        ///     LineSpacingBefore = "0";
        ///     LineSpacingAfter = "0";
        ///     Font = "Times New Roman";
        /// </summary>
        /// <param name="cellWidth"></param>
        /// <param name="justification"></param>
        /// <param name="fontSize"></param>
        public CustomFormat(JustificationValues justification, float fontSize)
        {
            CellWidth = -1;
            JustificationValue = justification;
            TableVerticalAlignmentValue = TableVerticalAlignmentValues.Center;
            LineSpacing = "240";
            LineSpacingRule = LineSpacingRuleValues.Auto;
            LineSpacingBefore = "0";
            LineSpacingAfter = "0";
            FontSize = fontSize;
            Font = "Times New Roman";
        }
    }
}
