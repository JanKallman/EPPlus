using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfNumberFormat : DxfStyleBase<ExcelDxfNumberFormat>
    {
        public ExcelDxfNumberFormat(ExcelStyles styles) : base(styles)
        {

        }
        int _numFmtID=int.MinValue;
        /// <summary>
        /// Id for number format
        /// 
        /// Build in ID's
        /// 
        /// 0   General 
        /// 1   0 
        /// 2   0.00 
        /// 3   #,##0 
        /// 4   #,##0.00 
        /// 9   0% 
        /// 10  0.00% 
        /// 11  0.00E+00 
        /// 12  # ?/? 
        /// 13  # ??/?? 
        /// 14  mm-dd-yy 
        /// 15  d-mmm-yy 
        /// 16  d-mmm 
        /// 17  mmm-yy 
        /// 18  h:mm AM/PM 
        /// 19  h:mm:ss AM/PM 
        /// 20  h:mm 
        /// 21  h:mm:ss 
        /// 22  m/d/yy h:mm 
        /// 37  #,##0 ;(#,##0) 
        /// 38  #,##0 ;[Red](#,##0) 
        /// 39  #,##0.00;(#,##0.00) 
        /// 40  #,##0.00;[Red](#,##0.00) 
        /// 45  mm:ss 
        /// 46  [h]:mm:ss 
        /// 47  mmss.0 
        /// 48  ##0.0E+0 
        /// 49  @
        /// </summary>            
        public int NumFmtID 
        { 
            get
            {
                return _numFmtID;
            }
            internal set
            {
                _numFmtID = value;
            }
        }
        string _format="";
        public string Format
        { 
            get
            {
                return _format;
            }
            set
            {
                _format = value;
                NumFmtID = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
            }
        }

        protected internal override string Id
        {
            get
            {
                return Format;
            }
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (NumFmtID < 0 && !string.IsNullOrEmpty(Format))
            {
                NumFmtID = _styles._nextDfxNumFmtID++;
            }
            helper.CreateNode(path);
            SetValue(helper, path + "/@numFmtId", NumFmtID);
            SetValue(helper, path + "/@formatCode", Format);
        }
        protected internal override bool HasValue
        {
            get 
            { 
                return !string.IsNullOrEmpty(Format); 
            }
        }
        protected internal override ExcelDxfNumberFormat Clone()
        {
            return new ExcelDxfNumberFormat(_styles) { NumFmtID = NumFmtID, Format = Format };
        }
    }
}
