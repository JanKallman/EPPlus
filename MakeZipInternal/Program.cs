using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace MakeZipInternal
{
    class Program
    {
        static void Main(string[] args)
        {
            
            foreach(string f in Directory.GetFiles(@"..\..\..\epplus\packaging\Dotnetzip\", "*.cs"))
            {
                if (f.ToLower().IndexOf("exception") == -1)
                {
                    string text = File.ReadAllText(f);
                    text = text.Replace("public class", "internal class");
                    text = text.Replace("public partial ", "internal partial ");
                    text = text.Replace("public abstract ", "internal abstract ");
                    text = text.Replace("public enum", "internal enum");
                    text = text.Replace("public event", "internal event");
                    File.WriteAllText(f, text);
                }
            }
        }
    }
}
