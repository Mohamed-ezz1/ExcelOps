
using DataHandler;

namespace Persons.Models;
public class Person
{
    [ExcelColumn("Name")]
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Email { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public int Experience { get; set; }
    public int idea { get; set; }


}
