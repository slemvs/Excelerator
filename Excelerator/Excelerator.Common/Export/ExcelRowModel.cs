namespace Excelerator.Export
{
	public class ExcelRowModel<TModel>
		where TModel : class
	{
		public TModel Data { get; set; }
	}
}
