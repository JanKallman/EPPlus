/* Copyright (C) 2011  Jan Källman
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 6
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2018-1-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public class ParsedValue
    {
        public ParsedValue(object val, int rowIndex, int colIndex)
        {
            Value = val;
            RowIndex = rowIndex;
            ColIndex = colIndex;
        }

        public object Value
        {
            get;
            private set;
        }

        public int RowIndex
        {
            get;
            private set;
        }

        public int ColIndex
        {
            get;
            private set;
        }
    }
}
