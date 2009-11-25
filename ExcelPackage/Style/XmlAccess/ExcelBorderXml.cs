/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Style.XmlAccess
{
    public class ExcelBorderXml : StyleXmlHelper
    {
        internal ExcelBorderXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {

        }
        internal ExcelBorderXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _left = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(leftPath, nsm));
            _right = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(rightPath, nsm));
            _top = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(topPath, nsm));
            _bottom = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(bottomPath, nsm));
            _diagonal = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(diagonalPath, nsm));
        }
        internal override string Id
        {
            get
            {
                return Left.Id + Right.Id + Top.Id + Bottom.Id + Diagonal.Id + DiagonalUp.ToString() + DiagonalDown.ToString();
            }
        }
        const string leftPath = "d:left";
        ExcelBorderItemXml _left = null;
        public ExcelBorderItemXml Left
        {
            get
            {
                return _left;
            }
            internal set
            {
                _left = value;
            }
        }
        const string rightPath = "d:right";
        ExcelBorderItemXml _right = null;
        public ExcelBorderItemXml Right
        {
            get
            {
                return _right;
            }
            internal set
            {
                _right = value;
            }
        }
        const string topPath = "d:top";
        ExcelBorderItemXml _top = null;
        public ExcelBorderItemXml Top
        {
            get
            {
                return _top;
            }
            internal set
            {
                _top = value;
            }
        }
        const string bottomPath = "d:bottom";
        ExcelBorderItemXml _bottom = null;
        public ExcelBorderItemXml Bottom
        {
            get
            {
                return _bottom;
            }
            internal set
            {
                _bottom = value;
            }
        }
        const string diagonalPath = "d:diagonal";
        ExcelBorderItemXml _diagonal = null;
        public ExcelBorderItemXml Diagonal
        {
            get
            {
                return _diagonal;
            }
            internal set
            {
                _diagonal = value;
            }
        }
        const string diagonalUpPath = "@diagonalUp";
        bool _diagonalUp = false;
        public bool DiagonalUp
        {
            get
            {
                return _diagonalUp;
            }
            internal set
            {
                _diagonalUp = value;
            }
        }
        const string diagonalDownPath = "@diagonalDown";
        bool _diagonalDown = false;
        public bool DiagonalDown
        {
            get
            {
                return _diagonalDown;
            }
            internal set
            {
                _diagonalDown = value;
            }
        }

        internal ExcelBorderXml Copy()
        {
            ExcelBorderXml newBorder = new ExcelBorderXml(NameSpaceManager);
            newBorder.Bottom = _bottom.Copy();
            newBorder.Diagonal = _diagonal.Copy();
            newBorder.Left = _left.Copy();
            newBorder.Right = _right.Copy();
            newBorder.Top = _top.Copy();
            newBorder.DiagonalUp = _diagonalUp;
            newBorder.DiagonalDown = _diagonalDown;

            return newBorder;

        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            CreateNode(leftPath);
            topNode.AppendChild(_left.CreateXmlNode(TopNode.SelectSingleNode(leftPath, NameSpaceManager)));
            CreateNode(rightPath);
            topNode.AppendChild(_right.CreateXmlNode(TopNode.SelectSingleNode(rightPath, NameSpaceManager)));
            CreateNode(topPath);
            topNode.AppendChild(_top.CreateXmlNode(TopNode.SelectSingleNode(topPath, NameSpaceManager)));
            CreateNode(bottomPath);
            topNode.AppendChild(_bottom.CreateXmlNode(TopNode.SelectSingleNode(bottomPath, NameSpaceManager)));
            CreateNode(diagonalPath);
            topNode.AppendChild(_diagonal.CreateXmlNode(TopNode.SelectSingleNode(diagonalPath, NameSpaceManager)));
            if (_diagonalUp)
            {
                SetXmlNode(diagonalUpPath, "1");
            }
            if (_diagonalDown)
            {
                SetXmlNode(diagonalDownPath, "1");
            }
            return topNode;
        }
    }
}
