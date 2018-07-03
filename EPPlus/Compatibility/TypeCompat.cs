/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

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
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Added       		        2017-11-02
 *******************************************************************************/
 using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
namespace OfficeOpenXml.Compatibility
{
    internal class TypeCompat
    {
        public static bool IsPrimitive(object v)
        {
#if (Core)            
            return v.GetType().GetTypeInfo().IsPrimitive;
#else
            return v.GetType().IsPrimitive;
#endif
        }
        public static bool IsSubclassOf(Type t, Type c)
        {
#if (Core)            
            return t.GetTypeInfo().IsSubclassOf(c);
#else
            return t.IsSubclassOf(c);
#endif
        }

        internal static bool IsGenericType(Type t)
        {
#if (Core)            
            return t.GetTypeInfo().IsGenericType;
#else
            return t.IsGenericType;
#endif

        }
        public static object GetPropertyValue(object v, string name)
        {
#if (Core)
            return v.GetType().GetTypeInfo().GetProperty(name).GetValue(v, null);
#else
            return v.GetType().GetProperty(name).GetValue(v, null);
#endif
        }
    }
}
