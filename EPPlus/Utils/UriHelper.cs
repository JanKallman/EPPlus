using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal class UriHelper
    {
        internal static Uri ResolvePartUri(Uri sourceUri, Uri targetUri)
        {
           if(targetUri.OriginalString.StartsWith("/"))
            {
                return targetUri;
            }
            string[] source = sourceUri.OriginalString.Split('/');
            string[] target = targetUri.OriginalString.Split('/');

            int t = target.Length - 1;
            int s;
            if(sourceUri.OriginalString.EndsWith("/")) //is the source a directory?
            {
                s = source.Length-1;
            }
            else
            {
                s=source.Length-2;
            }

            string file = target[t--];

            while (t >= 0)
            {
                if (target[t] == ".")
                {
                    break;
                }
                else if (target[t] == "..")
                {
                    s--;
                    t--;
                }
                else
                {
                    file = target[t--] + "/" + file;
                }
            }
            if (s >= 0)
            {
                for(int i=s;i>=0;i--)
                {
                    file = source[i] + "/" + file;
                }
            }
            return new Uri(file,UriKind.RelativeOrAbsolute);
        }

        internal static Uri GetRelativeUri(Uri WorksheetUri, Uri uri)
        {
            string[] source = WorksheetUri.OriginalString.Split('/');
            string[] target = uri.OriginalString.Split('/');

            int slen;
            if (WorksheetUri.OriginalString.EndsWith("/"))
            {
                slen = source.Length;
            }
            else
            {
                slen = source.Length-1;
            }
            int i = 0;
            while (i < slen && i < target.Length && source[i] == target[i])
            {
                i++;
            }

            string dirUp="";
            for (int s = i; s < slen; s++)
            {
                dirUp += "../";
            }
            string file = "";
            for (int t = i; t < target.Length; t++)
            {                
                file += (file==""?"":"/") + target[t];
            }
            return new Uri(dirUp+file,UriKind.Relative);
        }
    }
}
