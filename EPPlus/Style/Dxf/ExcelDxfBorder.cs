using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfBorderBase : DxfStyleBase<ExcelDxfBorderBase>
    {
        internal ExcelDxfBorderBase(ExcelStyles styles)
            : base(styles)
        {
            Left=new ExcelDxfBorderItem(_styles);
            Right = new ExcelDxfBorderItem(_styles);
            Top = new ExcelDxfBorderItem(_styles);
            Bottom = new ExcelDxfBorderItem(_styles);
        }
        /// <summary>
        /// Left border style
        /// </summary>
        public ExcelDxfBorderItem Left
        {
            get;
            internal set;
        }
        /// <summary>
        /// Right border style
        /// </summary>
        public ExcelDxfBorderItem Right
        {
            get;
            internal set;
        }
        /// <summary>
        /// Top border style
        /// </summary>
        public ExcelDxfBorderItem Top
        {
            get;
            internal set;
        }
        /// <summary>
        /// Bottom border style
        /// </summary>
        public ExcelDxfBorderItem Bottom
        {
            get;
            internal set;
        }
        ///// <summary>
        ///// Diagonal border style
        ///// </summary>
        //public ExcelDxfBorderItem Diagonal
        //{
        //    get;
        //    private set;
        //}
        ///// <summary>
        ///// A diagonal from the bottom left to top right of the cell
        ///// </summary>
        //public bool DiagonalUp
        //{
        //    get;
        //    set;
        //}
        ///// <summary>
        ///// A diagonal from the top left to bottom right of the cell
        ///// </summary>
        //public bool DiagonalDown
        //{
        //    get;
        //    set;
        //}

        protected internal override string Id
        {
            get
            {
                return Top.Id + Bottom.Id + Left.Id + Right.Id/* + Diagonal.Id + GetAsString(DiagonalUp) + GetAsString(DiagonalDown)*/;
            }
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            Left.CreateNodes(helper, path + "/d:left");
            Right.CreateNodes(helper, path + "/d:right");
            Top.CreateNodes(helper, path + "/d:top");
            Bottom.CreateNodes(helper, path + "/d:bottom");
        }
        protected internal override bool HasValue
        {
            get 
            {
                return Left.HasValue ||
                    Right.HasValue ||
                    Top.HasValue ||
                    Bottom.HasValue;
            }
        }
        protected internal override ExcelDxfBorderBase Clone()
        {
            return new ExcelDxfBorderBase(_styles) { Bottom = Bottom.Clone(), Top=Top.Clone(), Left=Left.Clone(), Right=Right.Clone() };
        }
    }
}
