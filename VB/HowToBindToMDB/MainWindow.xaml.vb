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
            Me.pivotGridControl1.BeginUpdate()
            ' Create a row pivot grid field bound to the Country datasource field.
            Dim fieldCountry As PivotGridField = New PivotGridField("Country", FieldArea.RowArea)
            ' Create a row pivot grid field bound to the Sales Person datasource field.
            Dim fieldCustomer As PivotGridField = New PivotGridField("Sales Person", FieldArea.RowArea)
            fieldCustomer.Caption = "Customer"
            ' Create a column pivot grid field bound to the OrderDate datasource field.
            Dim fieldYear As PivotGridField = New PivotGridField("OrderDate", FieldArea.ColumnArea)
            fieldYear.Caption = "Year"
            ' Group field values by years.
            fieldYear.GroupInterval = FieldGroupInterval.DateYear
            ' Create a column pivot grid field bound to the CategoryName datasource field.
            Dim fieldCategoryName As PivotGridField = New PivotGridField("CategoryName", FieldArea.ColumnArea)
            fieldCategoryName.Caption = "Product Category"
            ' Create a filter pivot grid field bound to the ProductName datasource field.
            Dim fieldProductName As PivotGridField = New PivotGridField("ProductName", FieldArea.FilterArea)
            fieldProductName.Caption = "Product Name"
            ' Create a data pivot grid field bound to the 'Extended Price' datasource field.
            Dim fieldExtendedPrice As PivotGridField = New PivotGridField("Extended Price", FieldArea.DataArea)
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
