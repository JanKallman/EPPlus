using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class TempFile
{
    public static TempFile Default { get; }

    public static TempFile GenerateFile
    {
        get { return Default.WithFile(); }
    }
    static Random Random = new Random((int)DateTime.UtcNow.Ticks);
    static string StringSet = "AbqQbCdEfhDiaSdFasdferTydFx";
    
    static IEnumerable<char> GenerateChars(int max=10){
        var num = Random.Next();
        while(max-- > 0){
            num = (num/StringSet.Length);
            yield return StringSet[num % StringSet.Length];	
            if (num == 0) num = Random.Next();
        }
    }
    static string GenerateString(int max = 10) {
        return new string(GenerateChars(max).ToArray());
    }

    static TempFile()
    {
        Default = new TempFile(
                Path.Combine(Path.GetTempPath(), GenerateString(10) + "." + GenerateString(3)))
            .Directory();
    }

    public TempFile Directory(string named=null)
    {
        var root = Path.Combine(this.DirectoryName, named ?? GenerateString());
        return new TempFile(Path.Combine(root, this.FileName));
    }

    public TempFile WithFile(string named = null)
    {
        return new TempFile(Path.Combine(this.DirectoryName, named ?? GenerateString(10) + "." + GenerateString(3)));
    }

    public TempFile WithFileExtention(string named)
    {
        return new TempFile(Path.ChangeExtension(this.FullPath, named));
    }

    public string FullPath { get; }

    public string DirectoryName
    {
        get
        {
            return Path.GetDirectoryName(FullPath);
        }
    }

    public string FileName
    {
        get { return Path.GetFileName(FullPath); }
    }
    public TempFile(string path)
    {
        System.IO.Directory.CreateDirectory(Path.GetDirectoryName(path));
        FullPath = path;
    }

    public static implicit operator FileInfo(TempFile file)
    {
        return new FileInfo(file.FullPath);
    }
}