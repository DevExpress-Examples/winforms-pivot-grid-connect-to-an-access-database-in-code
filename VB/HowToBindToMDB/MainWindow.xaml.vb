Imports System.Data
Imports System.Data.OleDb
Imports System.Windows
Imports DevExpress.Xpf.PivotGrid

Namespace HowToBindToMDB

	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			' Create a connection object.
			Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=NWIND.MDB")
			' Create a data adapter.
			Dim adapter As New OleDbDataAdapter("SELECT * FROM SalesPerson", connection)

			' Create and fill a dataset.
			Dim sourceDataSet As New DataSet()
			adapter.Fill(sourceDataSet, "SalesPerson")

			' Assign the data source to the PivotGrid control.
			pivotGridControl1.DataSource = sourceDataSet.Tables("SalesPerson")
			pivotGridControl1.DataProcessingEngine = DataProcessingEngine.Optimized

			pivotGridControl1.BeginUpdate()

			' Create a row pivot grid field bound to the Country data source column.
			Dim fieldCountry As New PivotGridField()
			fieldCountry.Caption = "Country"
			fieldCountry.Area = FieldArea.RowArea

			Dim countryBinding As New DataSourceColumnBinding("Country")
			fieldCountry.DataBinding = countryBinding

			' Create a row pivot grid field bound to the Sales Person data source column.
			Dim fieldCustomer As New PivotGridField()
			fieldCustomer.Caption = "Customer"
			fieldCustomer.Area = FieldArea.RowArea

			Dim customerBinding As New DataSourceColumnBinding("Sales Person")
			fieldCustomer.DataBinding = customerBinding

			' Create a column pivot grid field bound to the OrderDate data source column.
			Dim fieldYear As New PivotGridField()
			fieldYear.Caption = "Year"
			fieldYear.Area = FieldArea.ColumnArea

			Dim fieldOrderDate1Binding As New DataSourceColumnBinding("OrderDate")
			fieldOrderDate1Binding.GroupInterval = FieldGroupInterval.DateYear
			fieldYear.DataBinding = fieldOrderDate1Binding

			' Create a column pivot grid field bound to the CategoryName data source column.
			Dim fieldCategoryName As New PivotGridField()
			fieldCategoryName.Caption = "Product Category"
			fieldCategoryName.Area = FieldArea.ColumnArea

			Dim categoryNameBinding As New DataSourceColumnBinding("CategoryName")
			fieldCategoryName.DataBinding = categoryNameBinding

			' Create a filter pivot grid field bound to the ProductName data source column.
			Dim fieldProductName As New PivotGridField()
			fieldProductName.Caption = "Product Name"
			fieldProductName.Area = FieldArea.FilterArea

			Dim productNameBinding As New DataSourceColumnBinding("ProductName")
			fieldProductName.DataBinding = productNameBinding

			' Create a data pivot grid field bound to the 'Extended Price' data source column.
			Dim fieldExtendedPrice As New PivotGridField()
			fieldExtendedPrice.Area = FieldArea.DataArea

			Dim extendedPriceBinding As New DataSourceColumnBinding("Extended Price")
			fieldExtendedPrice.DataBinding = extendedPriceBinding

			' Specify the formatting setting to format summary values as integer currency amount.
			fieldExtendedPrice.CellFormat = "c0"

			' Add the fields to the control's field collection.         
			pivotGridControl1.Fields.AddRange(fieldCountry, fieldCustomer, fieldCategoryName, fieldProductName, fieldYear, fieldExtendedPrice)

			' Arrange the row fields within the Row Header Area.
			fieldCountry.AreaIndex = 0
			fieldCustomer.AreaIndex = 1

			' Arrange the column fields within the Column Header Area.
			fieldCategoryName.AreaIndex = 0
			fieldYear.AreaIndex = 1

			pivotGridControl1.EndUpdate()
		End Sub
	End Class
End Namespace
