using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;

namespace RCG
{
    public abstract class BaseFormatter : IFormatter
    {
        public GenProcessor Engine { get; protected set; }
        protected string _formatString = string.Empty;

        public Color ForeColor { get; protected set; }
        public Color BackColor { get; protected set; }
        public bool FontBold { get; protected set; }
        public bool FontItalic { get; protected set; }

        protected BaseFormatter(GenProcessor engine)
        {
            this.Engine = engine;
            ResetDefaultFormats();
        }

        protected void ResetDefaultFormats()
        {
            this.FormaterType = FormatTypes.None;
            this.BackColor = Color.White;
            this.ForeColor = Color.Black;
            this.FontBold = false;
            this.FontItalic = false;
        }

        protected void ParseFormatString(string formatString)
        {
            ResetDefaultFormats();
            string[] colorSettings = formatString.Split(';');
            foreach (string cs in colorSettings)
            {
                string colorPattern = string.Empty;
                string colorValue = string.Empty;

                if (!cs.Contains(':'))
                {
                    colorPattern = cs;
                }
                else
                {
                    colorPattern = cs.Split(':')[0];
                    colorValue = cs.Split(':')[1];
                }
                
                switch (colorPattern)
                {
                    case "fore-color":
                        this.FormaterType = this.FormaterType | FormatTypes.ForeColor;
                        ForeColor = Color.FromName(colorValue);
                        break;
                    case "back-color":
                        this.FormaterType = this.FormaterType | FormatTypes.BackColor;
                        BackColor = Color.FromName(colorValue);
                        break;
                    case "font-bold":
                        this.FormaterType = this.FormaterType | FormatTypes.FontBold;
                        FontBold = true;
                        break;
                    case "font-italic":
                        this.FormaterType = this.FormaterType | FormatTypes.FontItalic;
                        FontItalic = true;
                        break;
                    default:
                        throw new ArgumentException(string.Format("Not support format style {0} ...", colorPattern));
                }
            }
        }

        #region IFormatter Members

        public string Rule { get; set; }

        public string FormatString
        {
            get { return _formatString; }
            set
            {
                _formatString = value;
                ParseFormatString(value);
            }
        }

        public FormatTypes FormaterType { get; set; }

        public abstract bool Match(DataRow dr, FormatterConfig formatterConfig);

        public virtual void Execute(int currentExcelRowIndex)
        {
            ExcelOperationWrapper.ClearRowFormats(Engine.CurrentActiveExcelSheet, currentExcelRowIndex);

            if (this.FormaterType.HasFlag(FormatTypes.BackColor))
                ExcelOperationWrapper.SetRowBackgroundColor(Engine.CurrentActiveExcelSheet, currentExcelRowIndex, this.BackColor);
            if (this.FormaterType.HasFlag(FormatTypes.ForeColor))
                ExcelOperationWrapper.SetRowForegroundColor(Engine.CurrentActiveExcelSheet, currentExcelRowIndex, this.ForeColor);
            if (this.FormaterType.HasFlag(FormatTypes.FontBold))
                ExcelOperationWrapper.SetRowFontBold(Engine.CurrentActiveExcelSheet, currentExcelRowIndex, true);
            if (this.FormaterType.HasFlag(FormatTypes.FontItalic))
                ExcelOperationWrapper.SetRowFontItalic(Engine.CurrentActiveExcelSheet, currentExcelRowIndex, true);
        }

        public override string ToString()
        {
            return string.Format("Type: {0}, FormatType: {1}, FormatString: {2}", this.GetType().FullName, this.FormaterType.ToString(), this.FormatString);
        }
        #endregion
    }
}
