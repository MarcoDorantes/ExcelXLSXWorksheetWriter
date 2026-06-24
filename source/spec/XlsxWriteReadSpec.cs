namespace spec;

using System.IO;
using System.Linq;
using System.Collections.Generic;

record Entry
{
    public int ID {get;set;}
    public string Name {get;set;}
}

[TestClass]
public sealed class XlsxWriteReadSpec
{
    [TestMethod]
    public void InMemoryXLSX()
    {
        List<nutility.ExcelPage> pages = [];
        nutility.ExcelPage p = new("tab1");
        pages.Add(p);
        p.AddColumn("ID");
        p.AddColumn("Name");
        p.AddRow([123,"name1"]);
        p.AddRow([321,"name2"]);
        byte[] bytes = null;
        using(MemoryStream stream1 = new())
        {
            nutility.ExcelDoc.SaveToXLSX(stream1, [.. pages]);
            bytes = stream1.ToArray();
        }

        using MemoryStream stream2 = new(bytes);
        using nutility.ExcelXLSXWorksheetReader xlsx = new(stream2);
        var tabs = string.Join('|', xlsx.GetTabNames());
        var read = xlsx.ReadTabObjects<Entry>("tab1").ToList();
        var data = string.Join('|', read.Select(b => $"{b.Instance.ID}:{b.Instance.Name}/"));

        Assert.IsFalse(read.Any(b => b.Errors != null), string.Join('|', read.Where(b => b.Errors != null).SelectMany(b => b.Errors)));
        Assert.AreEqual("tab1",tabs);
        Assert.IsNotNull(read);
        Assert.AreEqual("123:name1/|321:name2/",data);
    }
}