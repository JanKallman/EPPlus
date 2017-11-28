EPPlus 4.5 Beta 2

Visit https://github.com/JanKallman/EPPlus for the latest information

EPPlus-Create advanced Excel spreadsheet.

4.5.0.1 Beta 2
* Added sparkline support.
* Switched targetframework from netcoreapp2.0 to netstandardapp2.0
* Replaced CoreCompat.System.Drawing.v2 with System.Drawing.Common
* Fixed a few issues. See https://github.com/JanKallman/EPPlus/commits/master

4.5.0.0 Beta 1
* .Net Core support.
* Added ExcelPackage.Compatibility.IsWorksheets1Based to remove inconsistent 1 base collection on the worksheets collection.
  Note: .Net Core will have this property set to false, and .Net 3.5 and .Net 4 version will have this property set to true for backward compatibility reasons.
  This property can be set via the appsettings.json file in .Net Core or the app.config file. See sample project for examples.
* RoundedCorners property Add to ExcelChart
* DataTable propery Added  to ExcelPlotArea for charts
* Sort method added on ExcelRange
* Added functions NETWORKDAYS.INTL and NETWORKDAYS. 
* And a lot of bug fixes. See https://github.com/JanKallman/EPPlus/commits/master

4.1.1
* Fix VBA bug in Excel 2016 - 1708 and later

4.1
* Added functions Rank, Rank.eq, Rank.avg and Search
* Applied a whole bunch of pull requests...
 * Performance and memory usage tweeks
 * Ability to set and retrieve 'custom' extended application propeties.
 * Added style QuotePrefix
 * Added support for MajorTimeUnit and MinorTimeUnit to chart axes
 * Added GapWidth Property to BarChart and Gapwidth.
 * Added Fill and Border properties to ChartSerie.
 * Added support for MajorTimeUnit and MinorTimeUnit to chart axes
 * Insert/delete row/column now shifts named ranges, comments, tables and pivottables.
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.5 Fixes
* Switched to Visual Studio 2015 for code and sample projects.
* Added LineColor, MarkerSize, LineWidth and MarkerLineColor properties to line charts
* Added LineEnd properties to shapes
* Added functions Value, DateValue, TimeValue
* Removed WPF depedency.
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.4 Fixes
* Added functions Daverage, Dvar Dvarp, DMax, DMin DSum,  DGet, DCount and DCountA 
* Exposed the formula parser logging functionality via FormulaParserManager.
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.3 Fixes
* Added compilation directive for MONO (Thanks Danny)
* Added functions IfError, Char, Error.Type, Degrees, Fixed, IsNonText, IfNa and SumIfs
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.2 Fixes
* Fixes a whole bunch of bugs related to the cell store (Worksheet.InsertColumn, Worksheet.InsertRow, Worksheet.DeleteColumn, Worksheet.DeleteRow, Range.Copy, Range.Clear)
* Added functions Acos, Acosh, Asinh, Atanh, Atan, CountBlank, CountIfs, Mina, Offset, Median, Hyperlink, Rept
* Fix for reading Excel comment content from the t-element.
* Fix to make Range.LoadFromCollection work better with inheritence
* And alot of other small fixes

4.0.1 Fixes
* VBA unreadable content
* Fixed a few issues with InsertRow and DeleteRow
* Fixed bug in Average and AverageA 
* Handling of Div/0 in functions
* Fixed VBA CodeModule error when copying a worksheet.
* Value decoding when reading str element for cell value.
* Better exception when accessing a worksheet out of range in the Excelworksheets indexer.
* Added Small and Large function to formula parser. Performance fix when encountering an unknown function.
* Fixed handling strings in formulas
* Calculate hangs if formula start with a parenthes.
* Worksheet.Dimension returned an invalid range in some cases.
* Rowheight was wrong in some cases.
* ExcelSeries.Header had an incorrect validation check.

New features 4.0

Replaced Packaging API with DotNetZip
* This will remove any problems with Isolated Storage and enable multithreading


New Cell store
* Less memory consumption
* Insert columns (not on the range level)
* Faster row inserts,

Formula Parser
* Calculates all formulas in a workbook, a worksheet or in a specified range
* 100+ functions implemented
* Access via Calculate methods on Workbook, Worksheet and Range objects.
* Add custom/missing Excel functions via Workbook. FormulaParserManager.
* Samples added to the EPPlusSamples project.

The formula parser does not support Array Formulas
* Intersect operator (Space)
* References to external workbooks
* And probably a whole lot of other stuff as well :)

Performance
*Of course the performance of the formula parser is nowhere near Excels. Our focus has been functionality.

Agile Encryption (Office 2012-)
* Support for newer type of encryption.

Minor new features
* Chart worksheets
* New Chart Types Bubblecharts
* Radar Charts
* Area Charts
* And lots of bug fixes...

Beta 2 Changes
* Fixed bug when using RepeatColumns & RepeatRows at the same time.
* VBA project will be left untouched if it’s not accessed.
* Fixed problem with strings on save.
* Added locks to the cell store for access by multiple threads.
* Implemented Indirect function
* Used DisplayNameAttribute to generate column headers from LoadFromCollection
* Rewrote ExcelRangeBase.Copy function. 
* Added caching to Save ZipStream for Cells and shared strings to speed up the Save method.
* Added Missing InsertColumn and DeleteColumn
* Added pull request to support Date1904 
* Added pull request ExcelWorksheet. LoadFromDataReader

Release Candidate changes
* Fixed some problems with Range.Copy Function
* InsertColumn and Delete column didn't work in some cases
* Chart.DisplayBlankAs had the wrong default type in Excel 2010+
* Datavalidation list overflow caused corruption of the package
* Fixed a few Calculation when referring ranges (for example If function)
* Added ChartAxis.DisplayUnit
* Fixed a bug related to shared formulas
* Named styles failed in some cases.
* Style.Indent got an invalid value in some cases.
* Fixed a problem with AutofitColumns method.
* Performance fix.
* A whole lot of other small fixes.
