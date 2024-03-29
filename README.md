Export XLSX - Highcharts module
===============================

This plugin adds the capability to export from Highcharts to the Excel .xlsx format. This is an updated version of the original export-xlsx plugin. I removed the moment.js and jQuery requirements and switched from SheetJS to the open source [ExcelJS](https://github.com/exceljs/exceljs) library for exporting. ExcelJS provides many more options for custom export formatting than SheetJS.

### Demo

* https://jsfiddle.net/chrisvecchio/ds6ocjy0/

### Requirements

* Latest Highcharts (tested with 8.0.0), but should work with version 7.0.0+.
* Latest Highcharts [exporting](https://code.highcharts.com/modules/exporting.js) and [export-data](https://code.highcharts.com/modules/export-data.js) modules.
* [ExcelJS](https://github.com/exceljs/exceljs) version 3.5.0+.
* [FileSaver.js](https://github.com/eligrey/FileSaver.js) version 2.0.2+.
* [SSF.js](https://github.com/SheetJS/ssf/) version 0.10.2+. (only needed if using auto column width feature)

**Notes:**

- Plugin does NO verification of ExcelJS options specifications. I recommend [google](https://www.google.com/), [stackoverflow](https://stackoverflow.com/questions/tagged/exceljs) or the official ExcelJS [docs](https://github.com/exceljs/exceljs/) for ExcelJS questions.
- Plugin does not check for correctly specified Excel number/date formats. A description of Excel's number format codes by Microsoft can be found [here](https://support.office.com/en-us/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68) . There are plenty of other resources available on Excel number formats via a google search or stackoverflow.

### Installation

* Add `<script>` tag pointing to `export-xlsx.js` below the required sources listed above.

### Code

The latest code is available on github: [https://github.com/chris-vecchio/export-xlsx](https://github.com/chris-vecchio/export-xlsx)

### Usage

_The plugin adds the following new options to the native Highcharts_ `exporting` _options under the_ `xlsx` _key_.
#### worksheet
##### `autoFitColumns ` (Boolean)
Enable/disable auto column width calculation. Default: `false`
Auto column width calculation requires loading [SSF.js](https://github.com/SheetJS/ssf/) before export-xlsx. See the [demo](#-demo) for an example of how to load SSF.js.
##### `sheetName` (String)
Excel worksheet name (Excel restricts sheet name length to <= 31 characters) Default: `Sheet1`
##### `categoryColumn.title` (String)
Category column title in Excel. Default: `Highcharts default`
##### `categoryColumn.numberFormat` (String)
Category column Excel number format. Default: `undefined` for non-datetime x-axis, `yyyy-mm-dd hh:mm:ss` for datetime x-axis.
##### Experimental options
##### `headerStyle` (Object)
Column header row font and fill options. See [example options](#example-options) for supported attributes Default: `{}`


#### workbook
##### `fileProperties ` (Object)
Excel workbook file properties. List of [available properties](https://github.com/exceljs/exceljs/blob/master/lib/doc/workbook.js#L153). Default: `{}`

#### series
_The plugin adds the following new options to individual_ `series` _options under the_ `xlsx` _key_.
##### `name` (String)
Column title for series in exported Excel file Default: `series name`
##### `numberFormat` (String)
Series column Excel number format. Default: `undefined`


### Example options
```javascript
exporting: {
    filename: 'export_xlsx_example',
    xlsx: {
        worksheet: {
            autoFitColumns: true,
            sheetName: 'CustomSheetName',
            categoryColumn: {
                title: 'Date',
                numberFormat: 'yyyy-mm'
            },
            headerStyle: {
                font: {
                    color: '#FFFFFF',
                    bold: true
                },
                fill: {
                    color: '#414b56'
                }
            }
        },
        workbook: {
            fileProperties: {
                creator: "File Author",
                company: "File Company",
                created: new Date(2017, 11, 31)
            }
        }
    }
},
series: [{
    name: 'Less than High School',
    xlsx: {
        numberFormat: '0.0',
        name: 'Less than HS'
    },
    data: [6.9, 7.3, 6.9, 7.1, 6.6, 6.7, 7.3, 6.5, 7.6, 7.3, 7.7, 7.7, 7.7, 7.4, 8.4, 7.7, 8.1, 8.7, 8.6, 9.7, 9.8, 10.3, 10.8, 11.1, 12.4, 13.2, 14.0, 14.9, 15.2, 15.6, 15.3, 15.6, 14.9, 15.2, 14.7, 15.0, 15.3, 15.8, 14.9, 14.7, 14.6, 14.2, 13.5, 14.1, 15.6, 15.0, 15.4, 15.0, 14.3, 14.0, 14.1, 14.7, 14.5, 14.4, 14.5, 14.1, 14.3, 13.5, 12.8, 13.7, 13.0, 13.1, 12.8, 12.5, 12.9, 12.6, 12.4, 11.8, 11.7, 12.1, 12.0, 11.8, 12.0, 11.3, 11.1, 11.6, 11.0, 10.7, 10.8, 11.1, 10.5, 10.9, 10.7, 9.8, 9.4, 9.8, 9.4, 8.7, 9.2, 9.2, 9.5, 9.2, 8.5, 8.1, 8.6, 8.6, 8.3, 8.2, 8.6, 8.5, 8.7, 8.2, 8.3, 7.9, 7.9, 7.6, 6.8, 6.5, 7.1, 7.0, 7.4, 7.6, 7.5, 7.6, 6.4, 7.4, 8.5, 7.5, 7.8, 7.6, 7.4, 7.6, 6.6, 6.4, 6.3, 6.5, 7.0, 6.1, 6.7, 6.0, 5.2, 6.3, 5.5, 5.6, 5.6, 5.8, 5.5, 5.6, 5.0, 5.7, 5.6, 5.9, 5.6, 5.8, 5.7]
}]
```

# Changelog

## [1.05] - 2021-01-06

Export XLSX now uses [ExcelJS](https://github.com/exceljs/exceljs) instead of [SheetJS](https://github.com/SheetJS/sheetjs)

### Removed
- Moment.js and jQuery dependencies
### Added
- Now compatible with a custom [columnHeaderFormatter](https://api.highcharts.com/highcharts/exporting.csv.columnHeaderFormatter) function
- Correctly sets column headers for [HighCharts combination chart](https://www.highcharts.com/demo/combo)
- Hours, minutes, and seconds date parts are now correctly exported and formatted. Previously, only month, year, and day were exported for a dateTime xAxis. 
### Changed
- Minimum supported Highcharts version updated bumped to 7.0.0 from 6.1.0
- Rows are no longer filtered to visible data for Pie charts [Issue #2](https://github.com/chris-vecchio/export-xlsx-exceljs/issues/2)
- Date/time values are recognized as Excel dates and formatted correctly in the exported file
- Updated docs and added several examples to index.html
