using Autodesk.AutoCAD.DatabaseServices;
using System.Reflection;

var t = typeof(Cell);
Console.WriteLine($"Type={t.FullName}");
foreach (var p in t.GetProperties(BindingFlags.Public|BindingFlags.Instance).Where(p=>p.Name.Contains("DataLink")||p.Name.Contains("Link")))
{
    Console.WriteLine($"prop: {p.PropertyType.Name} {p.Name}");
}
foreach (var m in t.GetMethods(BindingFlags.Public|BindingFlags.Instance|BindingFlags.Static).Where(m=>m.Name.Contains("DataLink")||m.Name.Contains("Link")))
{
    Console.WriteLine($"method: {m}");
}
