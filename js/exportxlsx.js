(function(H) {
    var pick = H.pick;

    // This extends the getDataRows function to only include rows for points visible in
    // the current chart view. Author: TorsteinHonsi
    // Source: https://github.com/highcharts/highcharts/issues/7913#issuecomment-371052869
    H.wrap(H.Chart.prototype, 'getDataRows', function(proceed, multiLevelHeaders) {
        var rows = proceed.call(this, multiLevelHeaders),
            xMin = this.xAxis[0].min,
            xMax = this.xAxis[0].max;
        rows = rows.filter(function(row) {
            return typeof row.x !== 'number' || (row.x >= xMin && row.x <= xMax);
        });
        return rows;
    });

    // Add function to export data to XLSX file
    H.Chart.prototype.downloadXLSX = function() {
        var xlsxOptions = this.userOptions.exporting.xlsx;

        // Set cell_dates flag to true or options.exporting.xlsx.cellDates
        // The flag enables parsing of cell dates from the html table. Haven't
        // found a case where enabling this on non-date data makes a difference,
        // but I'm leaving it as an option for testing purposes.
        var cell_dates = xlsxOptions.cellDates ? xlsxOptions.cellDates : true;

        // Determine what dateformat will be shown in chart.getTable() by
        // parsing the first date cell. If no change to options.exporting.csv.dateFormat
        // then Highcharts uses strftime format %Y-%m-%d %H:%M:%S or YYYY-MM-DD HH:mm:ss
        // in javascript format.this.axes[0]
        var momentFormat = strfToMoment(this.options.exporting.csv.dateFormat);
        var parseDateFormat = moment.utc(this.getDataRows()[1][0], momentFormat).creationData().format;

        // Format for dates in exported Excel file
        var exportDateFormat = pick(xlsxOptions.worksheet.dateFormat, 'm/d/yyyy');

        // Set export worksheet name to options.exporting.xlsx.worksheet.name or a
        // default of 'Sheet1'. Excel worksheet name length cannot exceed 31 characters
        var wsName = pick(xlsxOptions.worksheet.name.substring(0, 31), 'Sheet1');

        // Parse chart's html data table
        var wb = XLSX.read(this.getTable(), {
            sheet: wsName,
            type: 'string',
            cellDates: cell_dates
        });

        var ws = wb.Sheets[wb.SheetNames[0]];

        // Store range of used worksheet rows/columns
        var wsUsedRange = XLSX.utils.decode_range(ws['!ref']);

        // Format chart category column if datetime xAxis
        if (this.axes[0].isDatetimeAxis) {
            // Set category column title if specified
            if (xlsxOptions.worksheet.categoryColumnTitle) {
                ws[XLSX.utils.encode_cell({r: 0, c: 0})].v = xlsxOptions.worksheet.categoryColumnTitle;
            }
            for (var row = wsUsedRange.s.r+1; row <= wsUsedRange.e.r; ++row) {
                var cell = ws[XLSX.utils.encode_cell({r: row, c: 0})];

                var cellDate = moment(cell.v, parseDateFormat, true);

                if (cellDate.isValid()) {
                    cell.t = "d";
                    cell.v = cellDate.toDate();
                    cell.z = exportDateFormat;
                    XLSX.utils.format_cell(cell);
                }
            }
        }
        // Apply user-defined number formats to cells in series columns. Series column
        // indexes start at 1 since the category/date column is first.
        for (var col = wsUsedRange.s.c+1; col <= wsUsedRange.e.c; ++col) {
            for (var row = wsUsedRange.s.r+1; row <= wsUsedRange.e.r; ++row) {
                var cell = ws[XLSX.utils.encode_cell({r: row, c: col})];

                if (!cell) continue;  // if cell doesn't exist, move to next

                // Format numeric cells only
                if (cell.t == 'n') {
                    var series = this.series[col-1];
                    var cell_z = pick(series.userOptions.xlsxFormat, false);
                    if (cell_z) {
                        cell.z = cell_z;
                    }
                }
            }
        }

        // If enabled, autofit columns by setting column widths to the width of the
        // cell with the most characters
        if (xlsxOptions.worksheet.autoFitColumns === true) {
            var ncols = wsUsedRange.e.c - wsUsedRange.s.c + 1;

            // Add a hidden html copy of the worksheet to the page so we can calculate
            // column widths.
            $('body').append(XLSX.utils.sheet_to_html(ws, {id: 'colwidthtable'}));
            $('#colwidthtable').css('display', 'none');
            var colwidths = [];

            for (var i = 1; i <= ncols; i++) {
                var columnvalues = [];

                $('#colwidthtable tbody tr td:nth-child(' + i + ')').each(function() {
                    columnvalues.push($(this).text());
                });

                // Determine the width of the longest cell in the column
                var maxlen = columnvalues.reduce(function(a, b) {return a.length > b.length ? a : b;}, '');

                // Add 1 character to column width to give the cell a bit of extra padding
                // Should we let user specify this option? Some may want more padding.
                colwidths.push(maxlen.length + 1);
            }

            // Create key/value object of calculated column widths recognized by js-xlsx
            // 'wch' means "character width" in js-xlsx terminology
            var wscols = [];
            for (var j = 0; j < colwidths.length; j++) {
                wscols.push({'wch': colwidths[j]});
            }
            ws['!cols'] = wscols;

            // Remove the hidden html table used for calculating column widths
            $('#colwidthtable').remove();
        }

        // Set any user specified workbook file properties. You can see the full list of
        // available properties at https://docs.sheetjs.com/#workbook-file-properties
        if (xlsxOptions.workbook.fileProperties) {
            wb.Props = {};

            $.each(xlsxOptions.workbook.fileProperties, function(prop, value) {
                // Date fix for Internet Explorer. See https://stackoverflow.com/a/11253436
                if (prop == 'CreatedDate') {
                    var createdDate = new Date(value.getFullYear(), value.getMonth(),
                                            value.getDate(), value.getHours(),
                                            value.getMinutes(), value.getSeconds())
                                            .toISOString().replace(/\.\d*/, "");
                    wb.Props.CreatedDate = createdDate;
                }
                else {
                    wb.Props[prop] = value;
                }
                
            });
        }

        XLSX.writeFile(wb, (pick(this.options.exporting.filename, this.getFilename())) + '.xlsx', {
            bookType: 'xlsx',
            cellDates: cell_dates,
        });
    };
}(Highcharts));

// Function to convert Highcharts strftime date formats to Moment.js formats
var replacements = {
    'a': 'ddd',
    'A': 'dddd',
    'b': 'MMM',
    'B': 'MMMM',
    'c': 'lll',
    'd': 'DD',
    '-d': 'D',
    'e': 'D',
    'F': 'YYYY-MM-DD',
    'H': 'HH',
    '-H': 'H',
    'I': 'hh',
    '-I': 'h',
    'j': 'DDDD',
    '-j': 'DDD',
    'k': 'H',
    'l': 'h',
    'm': 'MM',
    '-m': 'M',
    'M': 'mm',
    '-M': 'm',
    'p': 'A',
    'P': 'a',
    'S': 'ss',
    '-S': 's',
    'u': 'E',
    'w': 'd',
    'W': 'WW',
    'x': 'll',
    'X': 'LTS',
    'y': 'YY',
    'Y': 'YYYY',
    'z': 'ZZ',
    'Z': 'z',
    'f': 'SSS',
    '%': '%'
};

var strfToMoment = function(format) {
    // Break up format string based on strftime tokens
    var tokens = format.split(/(%\-?.)/);
    var momentFormat = tokens.map(function(token) {
        // Replace strftime tokens with moment format tokens
        if (token[0] === '%' && replacements.hasOwnProperty(token.substr(1))) {
            return replacements[token.substr(1)];
        }
        return token
    }).join('');
    return momentFormat;
};