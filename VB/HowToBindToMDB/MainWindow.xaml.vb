Imports System.Data
Imports System.Data.OleDb
Imports System.Windows
Imports DevExpress.Xpf.PivotGrid

Namespace HowToBindToMDB

    Public Partial Class MainWindow
        Inherits Window

        Public Sub New()
            Me.InitializeComponent()
        End Sub

        Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Create a connection object.
            Dim connection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=NWIND.MDB")
            ' Create a data adapter.
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM SalesPerson", connection)
            ' Create and fill a dataset.
            Dim sourceDataSet As DataSet = New DataSet()
            adapter.Fill(sourceDataSet, "SalesPerson")
            ' Assign the data source to the PivotGrid control.
            Me.pivotGridControl1.DataSource = sourceDataSet.Tables("SalesPerson")
            Me.pivotGridControl1.DataProcessingEngine = DataProcessingEngine.Optimized
            Me.pivotGridControl1.BeginUpdate()
            ' Create a row pivot grid field bound to the Country data source column.
            Dim fieldCountry As PivotGridField = New PivotGridField()
            fieldCountry.Caption = "Country"
            fieldCountry.Area = FieldArea.RowArea
            Dim countryBinding As DataSourceColumnBinding = New DataSourceColumnBinding("Country")
            fieldCountry.DataBinding = countryBinding
            ' Create a row pivot grid field bound to the Sales Person data source column.
            Dim fieldCustomer As PivotGridField = New PivotGridField()
            fieldCustomer.Caption = "Customer"
            fieldCustomer.Area = FieldArea.RowArea
            Dim customerBinding As DataSourceColumnBinding = New DataSourceColumnBinding("Sales Person")
            fieldCustomer.DataBinding = customerBinding
            ' Create a column pivot grid field bound to the OrderDate data source column.
            Dim fieldYear As PivotGridField = New PivotGridField()
            fieldYear.Caption = "Year"
            fieldYear.Area = FieldArea.ColumnArea
            Dim fieldOrderDate1Binding As DataSourceColumnBinding = New DataSourceColumnBinding("OrderDate")
            fieldOrderDate1Binding.GroupInterval = FieldGroupInterval.DateYear
            fieldYear.DataBinding = fieldOrderDate1Binding
            ' Create a column pivot grid field bound to the CategoryName data source column.
            Dim fieldCategoryName As PivotGridField = New PivotGridField()
            fieldCategoryName.Caption = "Product Category"
            fieldCategoryName.Area = FieldArea.ColumnArea
            Dim categoryNameBinding As DataSourceColumnBinding = New DataSourceColumnBinding("CategoryName")
            fieldCategoryName.DataBinding = categoryNameBinding
            ' Create a filter pivot grid field bound to the ProductName data source column.
            Dim fieldProductName As PivotGridField = New PivotGridField()
            fieldProductName.Caption = "Product Name"
            fieldProductName.Area = FieldArea.FilterArea
            Dim productNameBinding As DataSourceColumnBinding = New DataSourceColumnBinding("ProductName")
            fieldProductName.DataBinding = productNameBinding
            ' Create a data pivot grid field bound to the 'Extended Price' data source column.
            Dim fieldExtendedPrice As PivotGridField = New PivotGridField()
            fieldExtendedPrice.Area = FieldArea.DataArea
            Dim extendedPriceBinding As DataSourceColumnBinding = New DataSourceColumnBinding("Extended Price")
            fieldExtendedPrice.DataBinding = extendedPriceBinding
            ' Specify the formatting setting to format summary values as integer currency amount.
            fieldExtendedPrice.CellFormat = "c0"
            ' Add the fields to the control's field collection.         
            Me.pivotGridControl1.Fields.AddRange(fieldCountry, fieldCustomer, fieldCategoryName, fieldProductName, fieldYear, fieldExtendedPrice)
            ' Arrange the row fields within the Row Header Area.
            fieldCountry.AreaIndex = 0
            fieldCustomer.AreaIndex = 1
            ' Arrange the column fields within the Column Header Area.
            fieldCategoryName.AreaIndex = 0
            fieldYear.AreaIndex = 1
            Me.pivotGridControl1.EndUpdate()
        End Sub
    End Class
End Namespace
