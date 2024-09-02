<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128578388/24.2.1%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E2118)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Pivot Grid for WPF - Bind the Pivot Grid to an MS Access Database at Runtime

The following example demonstrates how to bind the Pivot Grid control to a `SalesPerson` view in the **nwind.mdb** database included with the installation.

<!-- default file list -->
## Files to Review

* [MainWindow.xaml](./CS/HowToBindToMDB/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/HowToBindToMDB/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/HowToBindToMDB/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/HowToBindToMDB/MainWindow.xaml.vb))
<!-- default file list end -->

## Example Overview

Follow the steps below to connect the Pivot Grid to a database.

1. Create an `OleDbConnection` object and specify the connection string in its constructor.
2. Create an `OleDbDataAdapter` instance to select records from the data source.
3. Create a new `DataSet` object and populate it with data.
4. Use the [PivotGridControl.DataSource](https://docs.devexpress.com/WPF/DevExpress.Xpf.PivotGrid.PivotGridControl.DataSource) property to assign the resulting data source to the Pivot Grid.

Follow the steps below to create and configure Pivot Grid fields.

1. Create a [PivotGridField](https://docs.devexpress.com/WPF/DevExpress.Xpf.PivotGrid.PivotGridField) object and add it to the [PivotGridControl.Fields](https://docs.devexpress.com/WPF/DevExpress.Xpf.PivotGrid.PivotGridControl.Fields) collection.
2. Specify the field’s area and position within this area. For this, use the [PivotGridFieldBase.Area](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPivotGrid.PivotGridFieldBase.Area) and [PivotGridField.AreaIndex](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPivotGrid.PivotGridFieldBase.AreaIndex) properties. AreaIndex can be set only after the field is added to the control’s field collection.
3. Create a [DataSourceColumnBinding](https://docs.devexpress.com/WPF/DevExpress.Xpf.PivotGrid.DataSourceColumnBinding) object for each field.
4. Set the [DataSourceColumnBindingBase.ColumnName](https://docs.devexpress.com/CoreLibraries/DevExpress.PivotGrid.DataBinding.DataSourceColumnBindingBase.ColumnName) property to the name of the column in the data source. The Pivot Grid fields obtain their values from columns in the data source.
5. Assign the `DataSourceColumnBinding` object to the field’s [PivotGridFieldBase.DataBinding](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPivotGrid.PivotGridFieldBase.DataBinding) property.

## Documentation

- [Optimized Calculation Engine](https://docs.devexpress.com/CoreLibraries/401367/devexpress-pivot-grid-core-library/pivot-grid-modes/in-memory-mode/pivot-grid-optimized-calculation-engine?v=22.1)
- [Binding to Data](https://docs.devexpress.com/WindowsForms/1842/controls-and-libraries/pivot-grid/binding-to-data)





<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=wpf-pivot-grid-connect-to-an-access-database-in-code&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=wpf-pivot-grid-connect-to-an-access-database-in-code&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
