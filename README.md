# 📄✨ Excel to C# Object Importer  

A lightweight .NET utility to **read Excel files (`.xlsx`)** and map data to **C# objects** using **ClosedXML**.  

## 🚀 Features  
✅ **Reads Excel files** and converts rows into C# objects  
✅ **Automatic type conversion** (`int`, `double`, `DateTime`, `string`)  
✅ **Two mapping methods**:  
   - 🔹 **By Property Names** (matches Excel headers to C# property names)  
   - 🔹 **By Custom Attributes** (define column mappings with attributes)  
✅ **Ignores empty cells** to prevent errors  

## 📌 Usage  

### 📌 1️⃣ Define a C# Model  
```csharp
public class Person  
{  
    public string Name { get; set; }  
    public int Age { get; set; }  
    public string Email { get; set; }  
}
```
📌 2️⃣ Read Excel Data into a List of Objects
```csharp
string filePath = @"C:\path\to\file.xlsx";
List<Person> people = ExcelImporter.ReadExcelDataUsingPropertyNames<Person>(filePath);
```
📌 3️⃣ (Optional) Use Attributes for Custom Column Names
```csharp
[ExcelColumn("Full Name")]
public string Name { get; set; }  
List<Person> people = ExcelImporter.ReadExcelDataUsingAttributes<Person>(filePath);
```

🔧 Requirements
💻 .NET 8+
📂 ClosedXML (installed automatically)


📜 License
📝 MIT License
