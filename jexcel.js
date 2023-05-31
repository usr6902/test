/******************************* TODOS
 * Range
 *  HorizontalAlignment
 *  VerticalAlignment
 *  WrapText
 *  MergeCells
 *  Formula
 *  
 * Rows
 *  RowHeight
 *  
 * Columns
 *  Columns("A:A").ColumnWidth
 *  
 * Font
 *  Size
 *  ColorIndex
 *  Name
 *   
****************************************/

(function (je) {

    /**
     * Hata mesajını yazar
     * @param {string} msg
     */
    function WriteError(msg) {

        var err = window.page && window.page.WriteError || window.DashBoardHelper && window.DashBoardHelper.WriteError || window.$Dashboard && window.$Dashboard.WriteError || window.$Msg && (window.$Msg.WriteError || window.$Msg.writeError);
        var self = page || this;
        if (err) {
            err.call(self, msg + "");
        }
        else {
            alert(msg);
        }
        throw msg;
    }

    //override ActiveXObject
    if (!isIE()) {
        window.ActiveXObject = function (progId) {
            if (progId == "Excel.Application")
                return new JExcel();
            WriteError("ActiveXObject bu tarayıcıda desteklenmiyor. Internet Explorer ile açmayı deneyin. Hata Program Kodu: " + progId);
        }
    }

    var jApi = {};

    (function (jApi) {

        //fora ilk aşamada tanımsız olduğundan sayfanın load olmasını bekliyoruz
        document.addEventListener("DOMContentLoaded", function (event) {
            if (!isIE()) {
                fora.ExtendClass(jApi.JWorkbook, ExcelJS.Workbook, {});
            }
        });

        var intercept = function (obj) {
            if (!obj || obj.__isProxy)
                return obj;
            return new Proxy(obj, {
                get: function (target, key, receiver) {
                    if (key === "__isProxy")
                        return true;
                    if (key === "autofit") {
                        return Reflect.apply(target["AutoFit"], receiver, []);
                    }
                    var p = getAllPropertyNames(target).find(function (k) { return convertToLowerCase(k, 1) === convertToPropKey(key) });
                    if (!p)
                        p = getAllPropertyNames(target).find(function (k) { return convertToLowerCase(k) === convertToLowerCase(key) });

                    if (p) {
                        key = p;
                    }

                    if (typeof target[key] === 'function') {
                        return new Proxy(target[key], {
                            apply: function (target, thisArg, argumentsList) {
                                var res = Reflect.apply(target, thisArg, argumentsList);
                                return res;
                            }
                        });
                    }
                    var res = Reflect.get(target, key);
                    return res;
                },
                set: function (target, prop, value) {
                    var p = getAllPropertyNames(target).find(function (k) { return convertToLowerCase(k, 1) === convertToPropKey(prop) });
                    if (!p)
                        p = getAllPropertyNames(target).find(function (k) { return convertToLowerCase(k) === convertToLowerCase(prop) });

                    if (p) {
                        prop = p;
                    }
                    return Reflect.set(target, prop, value);
                }
            })
        }

        var convertToLowerCase = function (val, startIndex) {
            if (!val)
                return val;
            if (startIndex === undefined)
                startIndex = 0;
            return !val.slice ? val : val.slice(0, startIndex) + val.slice(startIndex).toLowerCase();
        }

        var convertToPropKey = function (val) {
            if (!val)
                return val;
            return !val.slice ? val : val.slice(0, 1).toUpperCase() + val.slice(1).toLowerCase();
        }

        var getAllPropertyNames = function (obj) {
            return Object.keys(obj);

            ////todo: Object.keys() ile tüm propları alamadım. Bu da bazı caselerde kısır döngüye giriyor diye şimdilik keys ile devam
            //var props = [];
            //do {
            //    Object.getOwnPropertyNames(obj).forEach(function (prop) {
            //        if (props.indexOf(prop) === -1) {
            //            props.push(prop);
            //        }
            //    });
            //} while (obj = Object.getPrototypeOf(obj));

            //return props;
        }

        var _defineGetterAndSetter = function (obj, prop, getter, setter) {
            Object.defineProperty(obj, prop, {
                set: setter,
                get: getter,
                enumerable: true,
                configurable: true
            });
        }

        var _defineGetter = function (obj, prop, getter) {
            _defineGetterAndSetter(obj, prop, getter, undefined);
        }

        var _defineSetter = function (obj, prop, setter) {
            _defineGetterAndSetter(obj, prop, undefined, setter);
        }

        var _jColorsByIndex = function (index) {
            var colors = {
                "1": "000000",
                "2": "FFFFFF",
                "3": "FF0000",
                "4": "00FF00",
                "5": "0000FF",
                "6": "FFFF00",
                "7": "FF00FF",
                "8": "00FFFF",
                "9": "800000",
                "10": "008000",
                "11": "000080",
                "12": "808000",
                "13": "800080",
                "14": "008080",
                "15": "C0C0C0",
                "16": "808080",
                "17": "9999FF",
                "18": "993366",
                "19": "FFFFCC",
                "20": "CCFFFF",
                "21": "660066",
                "22": "FF8080",
                "23": "0066CC",
                "24": "CCCCFF",
                "25": "000080",
                "26": "FF00FF",
                "27": "FFFF00",
                "28": "00FFFF",
                "29": "800080",
                "30": "800000",
                "31": "008080",
                "32": "0000FF",
                "33": "00CCFF",
                "34": "CCFFFF",
                "35": "CCFFCC",
                "36": "FFFF99",
                "37": "99CCFF",
                "38": "FF99CC",
                "39": "CC99FF",
                "40": "FFCC99",
                "41": "3366FF",
                "42": "33CCCC",
                "43": "99CC00",
                "44": "FFCC00",
                "45": "FF9900",
                "46": "FF6600",
                "47": "666699",
                "48": "969696",
                "49": "003366",
                "50": "339966",
                "51": "003300",
                "52": "333300",
                "53": "993300",
                "54": "993366",
                "55": "333399",
                "56": "333333"
            }
            return { argb: colors[index + ""] };
        };

        jApi.JExcel = function (options) {
            if (isIE()) {
                var excelApp = {};
                excelApp = new ActiveXObject("Excel.Application");
                return excelApp;
            }
            var that = this;
            this.Caption = "J-Excel";
            this._workbooks = [];
            this._workbooks.Add = this._workbooks.add = function (opt) {
                opt = opt || {};
                opt.SheetsInNewWorkbook = opt.SheetsInNewWorkbook || that.SheetsInNewWorkbook || 1;
                that.ActiveWorkbook = new jApi.JWorkbook(opt);
                that._workbooks.push(that.ActiveWorkbook);
                return that.ActiveWorkbook;
            };
            this.Open = this._workbooks.Open = function (url, callback) {
                var xhr = new XMLHttpRequest();
                xhr.responseType = "blob";
                xhr.open("GET", url + "?t=" + new Date().getTime());
                xhr.onload = function () {
                    if (this.status >= 200 && this.status < 300) {
                        that.ActiveWorkbook.xlsx.load(xhr.response).then(function () {
                            var keys = Object.keys(that.ActiveWorkbook.Worksheets());
                            for (var i in keys) {
                                var sh = that.ActiveWorkbook.Worksheets(i);
                                that.ActiveSheet = intercept(toJWorksheet(sh, that.ActiveWorkbook));
                            }
                            callback(that.ActiveWorkbook);
                        })
                    } else {
                        WriteError(xhr.statusText);
                    }
                };
                xhr.onerror = function () {
                    WriteError(xhr.statusText);
                };
                xhr.send();
            }
            this.JsonToExcel = function (data) {

                if (fora.Inue(data)) {
                    WriteError("Excel'e aktarılacak veri bulunamadı.");
                }
                else if (!Array.isArray(data)) {
                    WriteError("Beklenmedik veri formatı. Format [...] ya da [[...],[...]] şeklinde olmalıdır.");
                }
                else if (data.length <= 0) {
                    WriteError("Excel'e aktarılacak veri bulunamadı.");
                }
                else if (!_checkAllItemTypeIsSame(data)) {
                    WriteError("Beklenmedik veri formatı. Format [...] ya da [[...],[...]] şeklinde olmalıdır.");
                }

                if (Array.isArray(data[0])) {
                    that.ActiveSheet.Delete();
                    for (var i = 0; i < data.length; i++) {
                        if (!_checkAllItemTypeIsSame(data[i])) {
                            WriteError("Beklenmedik veri formatı. Format [...] ya da [[...],[...]] şeklinde olmalıdır.");
                        }
                        var sheet = that.ActiveWorkbook.Worksheets.Add();
                        _dataToExcel(data[i], sheet);
                    }
                }
                else {
                    _dataToExcel(data, that.ActiveSheet);
                }
            };

            function _dataToExcel(data, sheet) {
                var cols = [];
                for (var i = 0; i < data.length; i++) {
                    for (var k in data[i]) {
                        if (cols.indexOf(k) == -1) {
                            cols.push(k);
                        }
                    }
                }

                sheet.columns = cols.map(function (v, i) {
                    return { key: v, header: fora.Translate(v) }
                });


                data.forEach(function (item, index) {
                    sheet.addRow(item)
                })

                sheet.getRow(1).font = { bold: true };
            }


            function _checkAllItemTypeIsSame(data) {
                if (fora.Inue(data) || !Array.isArray(data))
                    return false;
                if (data.length < 2)
                    return true;
                if (!(Array.isArray(data[0]) || data[0] instanceof Object))
                    WriteError("Beklenmedik tür: " + typeof data[0]);
                var first = Array.isArray(data[0]);
                for (var i = 1; i < data.length; i++) {
                    if (!(Array.isArray(data[i]) || data[i] instanceof Object))
                        WriteError("Beklenmedik tür: " + typeof data[i]);
                    if (first != Array.isArray(data[i]))
                        return false;
                }
                return true;
            }

            _defineGetterAndSetter(this, "ActiveSheet", function () { return this.ActiveWorkbook.ActiveWorkSheet }, function (value) { this.ActiveWorkbook.ActiveWorkSheet = value });

            this._workbooks.Add(options);
            _defineGetterAndSetter(this, "Workbooks", function () { return intercept(that._workbooks); }, function (value) { that._workbooks = value; });


            this.View = function () {
                var that = this;
                this.ActiveWorkbook.views = [{ activeTab: this.ActiveWorkbook.ActiveWorkSheet.orderNo - 1 }];
                this.ActiveWorkbook.SaveAs(this.Caption);
            }

            _defineSetter(this, "Visible", function (value) {
                if (value) {
                    this.ActiveWorkbook.views = [{ activeTab: this.ActiveWorkbook.ActiveWorkSheet.orderNo - 1 }];
                    this.ActiveWorkbook.SaveAs(this.Caption);
                };
            });
            //_defineSetter(this, "SheetsInNewWorkbook", function (value) { that._workbooks.Add({ SheetsInNewWorkbook: value }) });

            this.Quit = this.Close = function () {
                //nop
            };
            this.Worksheets = function (key) {
                return this.ActiveWorkbook.Worksheets(key);
            }
            this.FindFile = function (callback) {

                callback = callback || function () { }

                function showDialog() {
                    var elem = document.createElement("input");
                    elem.type = "file";

                    if (elem && document.createEvent) {
                        var evt = document.createEvent("MouseEvents");
                        evt.initEvent("click", true, false);
                        elem.dispatchEvent(evt);
                        elem.addEventListener('change', handleFileSelect, false);
                    }
                }

                function handleFileSelect(evt) {
                    var files = evt.target.files;
                    var f = files[0];
                    var fr = new FileReader();
                    fr.onload = function () {
                        var data = fr.result;
                        that.ActiveWorkbook.xlsx.load(data).then(function () {
                            var keys = Object.keys(that.ActiveWorkbook.Worksheets());
                            for (var i in keys) {
                                var sh = that.ActiveWorkbook.Worksheets(i);
                                that.ActiveSheet = intercept(toJWorksheet(sh, that.ActiveWorkbook))
                            }
                            callback();
                        })
                    };
                    fr.readAsArrayBuffer(f);
                }

                showDialog();
            };

            return intercept(this);
        };


        jApi.JWorkbook = function (options) {
            var that = this;
            this.ActiveWorkSheet = null;
            jApi.JWorkbook._Parent.constructor.call(this);
            var rws = this.removeWorksheet;
            this.removeWorksheet = function (wsid) {
                rws.apply(this, arguments);
                var keys = Object.keys(this._worksheets)
                this.ActiveWorkSheet = keys.length > 0 ? intercept(this._worksheets[keys[0]]) : null;
            }
            this._worksheets = [];
            for (i = 0; i < (options.SheetsInNewWorkbook || 1); i++) {
                var sh = new jApi.JWorksheet(this);
                this.ActiveWorkSheet = sh;
            }

            _defineGetterAndSetter(this, "ActiveSheet", function () { return this.ActiveWorkSheet }, function (value) { this.ActiveWorkSheet = value });
            this.Close = function () {
                //nop
            }
            this.Worksheets = function (key) {
                if (key === undefined || that._worksheets.length === 0)
                    return intercept(that._worksheets);
                return intercept(that._worksheets[key] || that._worksheets[Object.keys(that._worksheets)[0]]);
            }

            _defineGetter(this.Worksheets, "Count", function () { return Object.keys(that._worksheets).length });
            this.Worksheets.Add = this.AddWorkSheet = function (sheetName, options) {
                var sh = new jApi.JWorksheet(that, sheetName, options);
                that.ActiveWorkSheet = sh;
                return intercept(sh);
            }

            this.SaveAs = function (filePath) {
                if (that._isSaved === true)
                    return;
                var fileName = filePath.replace(/^.*[\\\/]/, '');
                if (fora.Inue(fileName))
                    WriteError("Geçersiz dosya adı: " + filePath);
                var ext = fileName.split('.').pop();
                if (fileName !== filePath)
                    fora.Alert("Bu versiyonda dizine dosya kaydetme işlemine izin verilmemektedir, dosya indirilecektir.");
                if (["xls", "xlsx"].indexOf(ext.toLowerCase()) == -1)
                    fileName += ".xlsx";
                that.xlsx.writeBuffer()
                    .then(function (buffer) { saveAs(new Blob([buffer]), fileName); that._isSaved = true; })
                    .catch(function (err) { WriteError('Error writing excel export: ' + err) });
            }
            return intercept(this);
        }


        jApi.JWorksheet = function (wb, sheetName, options) {
            var ws = wb.addWorksheet(sheetName, options);
            return intercept(toJWorksheet(ws, wb));
        }

        var toJWorksheet = function (ws, wb) {
            if (ws.__isJSheet)
                return ws;
            if (!ws.columns)
                ws.columns = [];
            ws.__columns = ws.columns;
            ws.__columns.AutoFit = function (obj) {
                var minimalWidth = 10;
                ws.__columns.forEach(function (column) {
                    var maxColumnLength = 0;
                    column.eachCell({ includeEmpty: true }, function (cell) {
                        maxColumnLength = Math.max(
                            maxColumnLength,
                            minimalWidth,
                            cell.value ? cell.value.toString().length : 0
                        );
                    });
                    column.width = maxColumnLength + 2;
                });
            };
            _defineGetter(ws.__columns, "Count", function () {
                return ws.columnCount;
            })

            //activex uyumluluğu için eklendi
            _defineGetter(ws, "UsedRange", function () { return ws; });

            _defineGetterAndSetter(ws, "Columns", function () { return intercept(ws.__columns); }, function (value) { ws.__columns = value; })

            _defineGetter(ws._rows, "Count", function () {
                return ws.rowCount;
            })
            _defineGetterAndSetter(ws, "Rows", function () { return intercept(ws._rows); }, function (value) { ws._rows = value; })

            ws.Cells = function (r, c, value) {

                var cell = new jApi.JCell(ws, r, c);
                if (arguments.length == 3) {
                    cell.Value = value;
                }
                return cell;
            }

            ws.Range = function (rangeCell, r2) {
                return new jApi.JRange(ws, rangeCell, r2);
            }

            ws.Delete = function () {
                wb.removeWorksheet(ws.id);
            }

            ws.Activate = function () {
                wb.ActiveWorkSheet = ws;
            }
            _defineGetterAndSetter(ws, "Name", function () { return ws.name }, function (value) { ws.name = value });

            _defineGetter(ws, "__isJSheet", function () { return true });
            return ws;
        }

        jApi.JRange = function (ws, rangeCell, r2) {
            rangeCell = rangeCell + ":" + r2;
            var startCell, endCell;
            var ranges = rangeCell.split(':');
            startCell = ranges.length > 0 ? ranges[0] : undefined;
            endCell = ranges.length > 1 ? ranges[1] : undefined;


            // Recalculate in case bottom left and top right are given
            if (endCell < startCell) {
                var temp = endCell
                endCell = startCell
                startCell = temp
            }

            var endCellColumn = endCell.match(/[a-z]+/gi) ? endCell.match(/[a-z]+/gi)[0] : undefined;
            var endRow = endCell.match(/[^a-z]+/gi) ? endCell.match(/[^a-z]+/gi)[0] : undefined;
            var startCellColumn = startCell.match(/[a-z]+/gi) ? startCell.match(/[a-z]+/gi)[0] : undefined;
            var startRow = startCell.match(/[^a-z]+/gi) ? startCell.match(/[^a-z]+/gi)[0] : undefined;

            endCellColumn = endCellColumn || "XFD";
            startCellColumn = startCellColumn || "A";

            // Recalculate in case bottom left and top right are given
            if (endCellColumn < startCellColumn) {
                var temp = endCellColumn
                endCellColumn = startCellColumn
                startCellColumn = temp
            }

            // Recalculate in case bottom left and top right are given
            if (endRow < startRow) {
                var temp = endRow
                endRow = startRow
                startRow = temp
            }

            var endColumn = ws.getColumn(endCellColumn)
            var startColumn = ws.getColumn(startCellColumn)

            if (!endColumn) WriteError('End column not found')
            if (!startColumn) WriteError('Start column not found')

            var endColumnNumber = endColumn.number
            var startColumnNumber = startColumn.number

            var cells = [];
            for (var y = parseInt(startRow); y <= parseInt(endRow); y++) {
                var row = ws.getRow(y)
                for (var x = startColumnNumber; x <= endColumnNumber; x++) {
                    var cell = row.getCell(x);
                    cells.push(cell);
                }
            }

            function eachCell(cb) {
                for (var cell in cells) {
                    cell = cells[cell];
                    cb(cell);
                }
            }

            this.Cells = cells;
            this.StartCell = startCellColumn + startRow;
            this.EndCell = endCellColumn + endRow
            this.Range = this.StartCell + ":" + this.EndCell;

            var fontChanged = function (v) {
                eachCell(function (cell) { cell.font = _font });
            }
            var borderChanged = function (v) {
                eachCell(function (cell) { cell.border = _border });
            }

            var interiorChanged = function () {
                eachCell(function (cell) { cell.fill = _interior });
            }
            var _alignment = {};
            var _font = new jApi.JFont(fontChanged);
            var _border = new jApi.JBorder(borderChanged);
            var _interior = new jApi.JInterior(interiorChanged);
            var _rowHeight = -1;
            var _columnWidth = -1;

            //var verticalAlignmentEnums = ['top', 'middle', 'bottom', 'distributed', 'justify'];           
            var verticalAlignmentEnums = { "-4107": "bottom", "-4108": "middle", "-4117": "distributed", "-4130": "justify", "-4160": "top" };

            //var horizontalAlignmentEnums = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
            var horizontalAlignmentEnums = { "-1408": "center", "7": "centerContinuous", "-4117": "distributed", "5": "fill", "1": "justify", "-4130": "justify", "-4131": "left", "-4152": "right" };
            //#region private getter and setters
            var mergeCellsSetter = function (value) {
                if (value === true)
                    ws.mergeCells(this.Range);
            }

            var vAlignmentGetter = function () {
                var keys = Object.keys(verticalAlignmentEnums);
                for (var n = 0; n < keys.length; n++) {
                    if (_alignment.vertical == verticalAlignmentEnums[keys[n]])
                        return keys[n];
                }
                return undefined;
            }

            var vAlignmentSetter = function (value) {
                _alignment.vertical = verticalAlignmentEnums[value] || "top";
                eachCell(function (cell) { cell.alignment = fora.Extend(cell.alignment, _alignment) });
            }

            var hAlignmentSetter = function (value) {
                _alignment.horizontal = horizontalAlignmentEnums[value] || "left";
                eachCell(function (cell) { cell.alignment = fora.Extend(cell.alignment, _alignment) });
            }

            var hAlignmentGetter = function () {
                var keys = Object.keys(horizontalAlignmentEnums);
                for (var n = 0; n < keys.length; n++) {
                    if (_alignment.horizontal == horizontalAlignmentEnums[keys[n]])
                        return keys[n];
                }
                return undefined;
            }

            var numberFormatGetter = function () {
                return _numberFormat;
            }

            var numberFormatSetter = function (value) {
                _numberFormat = value;
                eachCell(function (cell) { cell.numFmt = _numberFormat });
            }


            var fontGetter = function () {
                return _font;
            }

            var fontSetter = function (value) {
                _font = fora.Extend(_font, value);
                eachCell(function (cell) { cell.font = _font });
            }

            var bordersGetter = function () {
                return _border;
            }

            var bordersSetter = function (value) {
                _border = fora.Extend(_border, value);
                eachCell(function (cell) { cell.border = _border });
            }

            var interiorGetter = function () {
                return _interior;
            }

            var interiorSetter = function (value) {
                _interior = fora.Extend(_interior, value);
                eachCell(function (cell) { cell.fill = _interior });
            }

            var wrapTextGetter = function () {
                return _alignment.wrapText;
            }

            var wrapTextSetter = function (value) {
                _alignment.wrapText = value;
                eachCell(function (cell) { cell.alignment = _alignment });
            }

            var rowHeightGetter = function () {
                return _rowHeight;
            }

            var rowHeightSetter = function (value) {
                _rowHeight = value;
                for (var y = parseInt(startRow); y <= parseInt(endRow); y++) {
                    var row = this.getRow(y);
                    row.height = _rowHeight;
                }
            }

            var columnWidthGetter = function () {
                return _columnWidth;
            }

            var columnWidthSetter = function (value) {
                _columnWidth = value;
                for (var x = startColumnNumber; x <= endColumnNumber; x++) {
                    var col = ws.getColumn(x);
                    col.width = _columnWidth;
                }
            }

            //#endregion

            this.Merge = function () {
                ws.mergeCells(this.Range);
            }

            this.EntireRow = intercept({
                AutoFit: function () {
                    return;
                    //todo: performans problemi var
                    _alignment.shrinkToFit = true;
                    eachCell(function (cell) { cell.alignment = fora.Extend(cell.alignment, _alignment) });
                }
            })

            this.EntireColumn = intercept({
                AutoFit: function () {
                    return;
                    //todo: performans problemi var
                    for (var x = startColumnNumber; x <= endColumnNumber; x++) {
                        var dataMax = 0;
                        var column = ws.getColumn(x);
                        for (var j = 1; j < column.values.length; j++) {
                            var columnLength = column.values[j].length;
                            if (columnLength > dataMax) {
                                dataMax = columnLength;
                            }
                        }
                        column.width = dataMax < 10 ? 10 : dataMax;
                    }
                    //_alignment.wrapText = true;
                    //eachCell(function (cell) { cell.alignment = fora.Extend(cell.alignment, _alignment) });
                }
            })

            _defineSetter(this, "MergeCells", mergeCellsSetter);
            _defineGetterAndSetter(this, "VerticalAlignment", vAlignmentGetter, vAlignmentSetter);
            _defineGetterAndSetter(this, "HorizontalAlignment", hAlignmentGetter, hAlignmentSetter);
            _defineGetterAndSetter(this, "NumberFormat", numberFormatGetter, numberFormatSetter);
            _defineGetterAndSetter(this, "Font", fontGetter, fontSetter);
            //_defineGetterAndSetter(this, "font", fontGetter, fontSetter);
            _defineGetterAndSetter(this, "WrapText", wrapTextGetter, wrapTextSetter);
            _defineGetterAndSetter(this.EntireRow, "RowHeight", rowHeightGetter, rowHeightSetter);
            _defineGetterAndSetter(this, "ColumnWidth", columnWidthGetter, columnWidthSetter);
            _defineGetterAndSetter(this, "Borders", bordersGetter, bordersSetter);
            _defineGetterAndSetter(this, "Interior", interiorGetter, interiorSetter);
            //_defineGetterAndSetter(this, "interior", interiorGetter, interiorSetter);
            return intercept(this);
        }


        jApi.JCell = function (ws, r, c) {
            if (!ws) WriteError('worksheet not found')
            if (!r) WriteError('row number not found')
            if (!c) WriteError('column number not found')

            var row = ws.getRow(r)
            var col = ws.getColumn(c)
            var cell = row.getCell(c);

            var fontChanged = function (v) {
                cell.font = _font;
            }
            var borderChanged = function (v) {
                cell.border = _border;
            }

            var interiorChanged = function () {
                cell.fill = _interior;
            }
            var _alignment = {};
            var _font = new jApi.JFont(fontChanged);
            var _border = new jApi.JBorder(borderChanged);
            var _interior = new jApi.JInterior(interiorChanged);
            var _rowHeight = -1;
            var _columnWidth = -1;

            var verticalAlignmentEnums = ['top', 'middle', 'bottom', 'distributed', 'justify'];
            var horizontalAlignmentEnums = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
            //#region private getter and setters

            var vAlignmentGetter = function () {
                return verticalAlignmentEnums.indexOf(_alignment.vertical);
            }

            var vAlignmentSetter = function (value) {
                _alignment.vertical = verticalAlignmentEnums[value];
                cell.alignment = fora.Extend(cell.alignment, _alignment);
            }

            var hAlignmentSetter = function (value) {
                _alignment.horizontal = horizontalAlignmentEnums[value];
                cell.alignment = fora.Extend(cell.alignment, _alignment);
            }

            var hAlignmentGetter = function () {
                return horizontalAlignmentEnums.indexOf(_alignment.horizontal);
            }

            var numberFormatGetter = function () {
                return _numberFormat;
            }

            var numberFormatSetter = function (value) {
                _numberFormat = value;
                cell.numFmt = _numberFormat;
            }


            var fontGetter = function () {
                return _font;
            }

            var fontSetter = function (value) {
                _font = fora.Extend(_font, value);
                cell.font = _font;
            }

            var bordersGetter = function () {
                return _border;
            }

            var bordersSetter = function (value) {
                _border = fora.Extend(_border, value);
                cell.border = _border;
            }

            var interiorGetter = function () {
                return _interior;
            }

            var interiorSetter = function (value) {
                _interior = fora.Extend(_interior, value);
                cell.fill = _interior;
            }

            var wrapTextGetter = function () {
                return _alignment.wrapText;
            }

            var wrapTextSetter = function (value) {
                _alignment.wrapText = value;
                cell.alignment = _alignment;
            }

            var rowHeightGetter = function () {
                return _rowHeight;
            }

            var rowHeightSetter = function (value) {
                _rowHeight = value;
                row.height = _rowHeight;
            }

            var columnWidthGetter = function () {
                return _columnWidth;
            }

            var columnWidthSetter = function (value) {
                _columnWidth = value;
                col.width = _columnWidth;
            }

            var valueGetter = function () {
                if (cell.value == null)
                    return undefined;
                return cell.value;
            }

            var valueSetter = function (value) {
                cell.value = value;
            }

            this.EntireRow = intercept({
                AutoFit: function () {
                    return;
                    //todo: performans problemi var
                    _alignment.shrinkToFit = true;
                    cell.alignment = fora.Extend(cell.alignment, _alignment);
                }
            })

            this.EntireColumn = intercept({
                AutoFit: function () {
                    return;
                    //todo: performans problemi var
                    var dataMax = 0;
                    for (var j = 1; j < col.values.length; j++) {
                        var columnLength = col.values[j].length;
                        if (columnLength > dataMax) {
                            dataMax = columnLength;
                        }
                    }
                    col.width = dataMax < 10 ? 10 : dataMax;
                }
            });

            _defineGetterAndSetter(this, "VerticalAlignment", vAlignmentGetter, vAlignmentSetter);
            _defineGetterAndSetter(this, "HorizontalAlignment", hAlignmentGetter, hAlignmentSetter);
            _defineGetterAndSetter(this, "NumberFormat", numberFormatGetter, numberFormatSetter);
            _defineGetterAndSetter(this, "Font", fontGetter, fontSetter);
            //_defineGetterAndSetter(this, "font", fontGetter, fontSetter);
            _defineGetterAndSetter(this, "WrapText", wrapTextGetter, wrapTextSetter);
            _defineGetterAndSetter(this.EntireRow, "RowHeight", rowHeightGetter, rowHeightSetter);
            _defineGetterAndSetter(this, "ColumnWidth", columnWidthGetter, columnWidthSetter);
            _defineGetterAndSetter(this, "Borders", bordersGetter, bordersSetter);
            _defineGetterAndSetter(this, "Interior", interiorGetter, interiorSetter);
            //_defineGetterAndSetter(this, "interior", interiorGetter, interiorSetter);
            _defineGetterAndSetter(this, "Value", valueGetter, valueSetter);
            return intercept(this);
        }

        jApi.JFont = function (fontChanged) {
            fontChanged = fontChanged || fora.EmptyFunction;
            var _font = {};

            function getterFactory(prop) {
                return function () {
                    return _font[prop];
                }
            }

            function setterFactory(prop) {
                return function (value) {
                    _font[prop] = value;
                    fontChanged(_font)
                }
            }

            function colorIndexSetter(value) {
                _font.colorIndex = value;
                _font.color = _jColorsByIndex(_font.colorIndex);
            }

            //_defineGetterAndSetter(this, "bold", getterFactory("bold"), setterFactory("bold"));
            _defineGetterAndSetter(this, "Bold", getterFactory("bold"), setterFactory("bold"));
            _defineGetterAndSetter(this, "Name", getterFactory("name"), setterFactory("name"));
            //_defineGetterAndSetter(this, "name", getterFactory("name"), setterFactory("name"));
            _defineGetterAndSetter(this, "Size", getterFactory("size"), setterFactory("size"));
            //_defineGetterAndSetter(this, "size", getterFactory("size"), setterFactory("size"));
            _defineGetterAndSetter(this, "ColorIndex", getterFactory("colorIndex"), colorIndexSetter);
            //_defineGetterAndSetter(this, "colorIndex", getterFactory("colorIndex"), colorIndexSetter);
            return intercept(this);
        }

        jApi.JBorder = function (callback) {
            callback = callback || fora.EmptyFunction;
            function borderStyleChanged(bs) {
                callback(_border);
            }
            var _border = {
                top: new jApi.JBorderStyle(borderStyleChanged),
                left: new jApi.JBorderStyle(borderStyleChanged),
                right: new jApi.JBorderStyle(borderStyleChanged),
                bottom: new jApi.JBorderStyle(borderStyleChanged)
            };
            var styles = ['dotted', 'thin', 'dashDot', 'hair', 'dashDotDot', 'slantDashDot', 'mediumDashed', 'mediumDashDotDot', 'mediumDashDot', 'medium', 'double', 'thick'];

            function colorIndexSetter(value) {
                var c = _jColorsByIndex(value) || { argb: 'FF000000' };
                _border.top.color = c;
                _border.bottom.color = c;
                _border.right.color = c;
                _border.left.color = c;
                callback(_border);
            }

            function lineStyleSetter(value) {
                var ls = value + "";
                if (!ls.match(/[^0-9]+/gi))
                    ls = styles[parseInt(ls)] || styles[1];
                _border.top.style = ls;
                _border.bottom.style = ls;
                _border.right.style = ls;
                _border.left.style = ls;
                callback(_border);
            }

            _defineSetter(_border, "ColorIndex", colorIndexSetter);
            //_defineSetter(_border, "colorIndex", colorIndexSetter);
            _defineSetter(_border, "LineStyle", lineStyleSetter);
            //_defineSetter(_border, "lineStyle", lineStyleSetter);
            return intercept(_border);
        }

        jApi.JBorderStyle = function (callback) {
            callback = callback || fora.EmptyFunction;
            var _bs = {};
            function getterFactory(prop) {
                return function () {
                    return _bs[prop];
                }
            }

            function setterFactory(prop) {
                return function (value) {
                    _bs[prop] = value;
                    callback(_bs)
                }
            }

            _defineGetterAndSetter(this, "style", getterFactory("style"), setterFactory("style"));
            _defineGetterAndSetter(this, "color", getterFactory("color"), setterFactory("color"));
            return intercept(this);
        }

        jApi.JInterior = function (callback) {
            callback = callback || fora.EmptyFunction;

            var _interior = {
                type: 'pattern',
                pattern: 'solid'
            };

            function colorIndexSetter(value) {
                var c = _jColorsByIndex(value) || { argb: 'FF000000' };
                _interior.fgColor = c;
                callback(_interior);
            }

            _defineSetter(_interior, "ColorIndex", colorIndexSetter);
            //_defineSetter(_interior, "colorIndex", colorIndexSetter);
            return intercept(_interior);
        }

    })(jApi)
    ////-------------------------------------------------------------------------------
    //register it globally;
    window["JExcel"] = jApi.JExcel;
})()
