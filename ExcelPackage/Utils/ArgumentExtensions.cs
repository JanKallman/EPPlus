/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
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
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-01-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// Extension methods for guarding
    /// </summary>
    public static class ArgumentExtensions
    {

        /// <summary>
        /// Throws an ArgumentNullException if argument is null
        /// </summary>
        /// <typeparam name="T">Argument type</typeparam>
        /// <param name="argument">Argument to check</param>
        /// <param name="argumentName">parameter/argument name</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void IsNotNull<T>(this IArgument<T> argument, string argumentName)
            where T : class
        {
            argumentName = string.IsNullOrEmpty(argumentName) ? "value" : argumentName;
            if (argument.Value == null)
            {
                throw new ArgumentNullException(argumentName);
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentNullException"/> if the string argument is null or empty
        /// </summary>
        /// <param name="argument">Argument to check</param>
        /// <param name="argumentName">parameter/argument name</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void IsNotNullOrEmpty(this IArgument<string> argument, string argumentName)
        {
            if (string.IsNullOrEmpty(argument.Value))
            {
                throw new ArgumentNullException(argumentName);
            }
        }

        /// <summary>
        /// Throws an ArgumentOutOfRangeException if the value of the argument is out of the supplied range
        /// </summary>
        /// <typeparam name="T">Type implementing <see cref="IComparable"/></typeparam>
        /// <param name="argument">The argument to check</param>
        /// <param name="min">Min value of the supplied range</param>
        /// <param name="max">Max value of the supplied range</param>
        /// <param name="argumentName">parameter/argument name</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static void IsInRange<T>(this IArgument<T> argument, T min, T max, string argumentName)
            where T : IComparable
        {
            if (!(argument.Value.CompareTo(min) >= 0 && argument.Value.CompareTo(max) <= 0))
            {
                throw new ArgumentOutOfRangeException(argumentName);
            }
        }
    }
}
