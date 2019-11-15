# Export XLSX - Highcharts module

This plugin adds the capability to export from Highcharts to the Excel .xlsx format.

### Demo

* https://jsfiddle.net/chrisvecchio/f2r7jqhz/

### Requirements

* Latest Highcharts (tested with 7.02), but should work with version 6.1.0+.
* Latest Highcharts [exporting]https://code.highcharts.com/modules/exporting.js and [export-data]https://code.highcharts.com/modules/export-data.js modules.
* [Moment.js](http://momentjs.com/) version 2.22.2+.
* [SheetJS js-xlsx](https://github.com/SheetJS/js-xlsx) version 0.14.0+.
**Note:** this plugin does NO verification of js-xlsx options specification. I recommend [google](https://www.google.com/), [stackoverflow](https://stackoverflow.com/questions/tagged/js-xlsx) or the official js-xlsx [docs](https://docs.sheetjs.com/) for js-xlsx questions.

### Installation

* Add `<script>` tag pointing to `exportxlsx.js` below the required sources listed above.

### Code

The latest code is available on github: [https://github.com/chris-vecchio/export-xlsx](https://github.com/chris-vecchio/export-xlsx)

### Usage

The plugin adds the following new options to the native Highcharts exporting options.

<table>
<thead>
<tr>
<th>Property</th>
<th>Type</th>
<th>Default</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td align="left">cellDates</td>
<td align="left">boolean</td>
<td align="left">true</td>
<td align="left">Check for and parse xAxis dates from chart HTML table</td>
</tr>
<tr>
<td colspan=4 style="font-weight: bold;">worksheet</td>
</tr>
<tr>
<td align="left">autoFitColumns</td>
<td align="left">boolean</td>
<td align="left">false</td>
<td align="left">Enable auto column width calculation</td>
</tr>
<tr>
<td align="left">categoryColumnTitle</td>
<td align="left">str</td>
<td align="left">undefined</td>
<td align="left">Category column title in Excel. Leave empty for Highcharts default</td>
</tr>
<tr>
<td align="left">dateFormat</td>
<td align="left">str</td>
<td align="left">m/d/yyyy</td>
<td align="left">Excel date format for category column in Excel</td>
</tr>
<tr>
<td align="left">name</td>
<td align="left">str</td>
<td align="left">Sheet1</td>
<td align="left">Excel worksheet name (Excel restricts name length to <= 31 characters)</td>
</tr>
<tr>
<td colspan=4 style="font-weight: bold;">workbook</td>
</tr>
<tr>
<td align="left">fileProperties</td>
<td align="left">object</td>
<td align="left">undefined</td>
<td align="left">Workbook file properties. See description of available properties <a href="https://docs.sheetjs.com/#workbook-file-properties">here.</a></td>
</tr>
</tbody>
</table>

You can specify Excel number formats for individual series by adding the ```xlsxFormat``` option to individual series options.

**Note:** the plugin does not check for correctly specified Excel number/date formats. See the [number format documentation](https://docs.sheetjs.com/#number-formats) from SheetJS (or google/stackoverflow) for questions about number formats.

Example options:
```javascript
exporting: {
    xlsx: {
        categoryColumnTitle: 'Month',
        dateFormat: 'yyyy-mm',
        worksheet: {
            name: 'CustomWorksheetName',
            autoFitColumns: true
        },
        workbook: {
            fileProperties: {
                Author: "File Author",
                Company: "File Company",
                CreatedDate: new Date(Date.now())
            }
        }
    }
},
series: [{
    name: 'Less than high school',
    xlsxFormat: '0.0',
    data: [6.9, 7.3, 6.9, 7.1, 6.6, 6.7, 7.3, 6.5, 7.6, 7.3, 7.7, 7.7, 7.7, 7.4, 8.4, 7.7, 8.1, 8.7, 8.6, 9.7, 9.8, 10.3, 10.8, 11.1, 12.4, 13.2, 14.0, 14.9, 15.2, 15.6, 15.3, 15.6, 14.9, 15.2, 14.7, 15.0, 15.3, 15.8, 14.9, 14.7, 14.6, 14.2, 13.5, 14.1, 15.6, 15.0, 15.4, 15.0, 14.3, 14.0, 14.1, 14.7, 14.5, 14.4, 14.5, 14.1, 14.3, 13.5, 12.8, 13.7, 13.0, 13.1, 12.8, 12.5, 12.9, 12.6, 12.4, 11.8, 11.7, 12.1, 12.0, 11.8, 12.0, 11.3, 11.1, 11.6, 11.0, 10.7, 10.8, 11.1, 10.5, 10.9, 10.7, 9.8, 9.4, 9.8, 9.4, 8.7, 9.2, 9.2, 9.5, 9.2, 8.5, 8.1, 8.6, 8.6, 8.3, 8.2, 8.6, 8.5, 8.7, 8.2, 8.3, 7.9, 7.9, 7.6, 6.8, 6.5, 7.1, 7.0, 7.4, 7.6, 7.5, 7.6, 6.4, 7.4, 8.5, 7.5, 7.8, 7.6, 7.4, 7.6, 6.6, 6.4, 6.3, 6.5, 7.0, 6.1, 6.7, 6.0, 5.2, 6.3, 5.5, 5.6, 5.6, 5.8, 5.5, 5.6, 5.0, 5.7, 5.6, 5.9, 5.6, 5.8, 5.7]
}]
```
