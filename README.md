# Excelerator
Set of classes for extremely simple generation plain Excel files

##Get started
You can start working with Excelerator by creating a class which will be a row model for your Excel document. Let's define simple row model class.

```csharp
public class Example1Model
{
  public string Prop1 { get; set; }
  public string Prop2 { get; set; }
}
```

Then you have to create your implementation of `IExcelGenerator<Example1Model>` by inherit it from `ClosedXmlExcelGenerator<Example1Model>`.

```csharp
public class Example1Generator : ClosedXmlExcelGenerator<Example1Model> { }
```

Next step you define metadata for each column of your Excel document.

```csharp
var wsMetadata = new WorksheetMetadata<Example1Model>
{
	Name = "Example1",
	StartColumn = 5,
	StartRow = 3,
	FormatAsTable = true,
	ColumnsMetadata = new List<ColumnMetadata<Example1Model>>
	{
		new ColumnMetadata<Example1Model> {Header = "Prop1", Value = _ => _.Prop1},
		new ColumnMetadata<Example1Model> {Header = "Prop2", Value = _ => _.Prop2}
	}
};
```

That's all!

You only have to call `Generate` method of your object of type `Example1Generator` with metadata and collection of ExcelRowModel of your row model objects `List<ExcelRowModel<Example1Model>>` as parameters. As result you will get a memory stream. You can save it to file with *.xlsx* extension or attach it to e-mail.

```csharp
var data = new List<ExcelRowModel<Example1Model>>()
{
  new ExcelRowModel<Example1Model>() { Data = new Example1Model() { Prop1 = "prop1_1", Prop2 = "prop2_1" }},
  new ExcelRowModel<Example1Model>() { Data = new Example1Model() { Prop1 = "prop1_2", Prop2 = "prop2_2" }},
  new ExcelRowModel<Example1Model>() { Data = new Example1Model() { Prop1 = "prop1_3", Prop2 = "prop2_3" }}
}

var generator = new Example1Generator();
var ms = generator.Generate(wsMetadata, data)
```

Result file will be like this:

| Prop1  | Prop2 |
| ------------- | ------------- |
| prop1_1  | prop2_1  |
| prop1_2  | prop2_2  |
| prop1_3  | prop2_3  |
