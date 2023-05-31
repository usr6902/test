
//-------------------------------------------------------------------------------
// Excele çıkacak kayıtlar için öncelikle bu functiondan obje oluşturulacak ve 
// heaeder, format ve rows alanları set edilecek
function ExcelApplication() {

    //var obj = DashBoardHelper.GetMessageDll();
    //alert(obj.GetUserCultureInfo());
    this.FormatType = {
        Number: "0.00",
        Date: "dd.mm.yyyy",
        Decimal: "0",
        Currency: isIE() ? "$" : "#,##0.00 ₺;[Red]-#,##0.00 ₺",
        String: "",
        Money: "#,###,###.00"
    };

    this.header = new Array();
    this.headerMemberName = new Array();
    this.format = new Array();
    this.rows = new Array();
    this.visibleColumns = new Array();
    this.rowCount = 0;
    this.AllColumnsVisible = true;
    this.ShowSumColumn = false;
    this.AddTopHeader = null;
    this.AddCell = null;
    this.AddCellXY = null;
    this.CustomizeHeader = null;
    this.AddSubHeader = null;

    this.ReturnExcelObject = function (value) {
        var object = new JExcel();
        return object;
    };

    //todo: bu metod kullanılmıyor. ne yapılacağıyla ilgili karar verilmeli. implemente mi edelim yoksa silelim mi?
    this.AddSheet = function (value) {
        this.AddNewSheet();
    };

    this.AddHeader = function (value) {
        this.header[this.header.length] = value;
    };

    this.AddHeaderMemberName = function (value) {
        this.headerMemberName[this.headerMemberName.length] = value;
    };

    this.AddFormat = function (value) {
        this.format[this.format.length] = value;
    };

    this.AddNewRow = function () {
        this.rows[this.rowCount] = new Array();
        this.rowCount++;
    };

    this.AddCell = function (value) {
        if (this.rows[0] == null) {
            this.AddNewRow();
        }
        this.rows[this.rowCount - 1][this.rows[this.rowCount - 1].length] = value;
    };

    this.SetHeader = function (arr) {
        this.header = arr;
    };


    //Excele çıkacak kayıtlar için Datagriddeki başlıkları alıyoruz.
    this.SetHeaderFromDataGrid = function (columns) {
        var header = new Array();
        var ind = 0;
        for (var prop in columns) {
            if (columns.hasOwnProperty(prop)) {
                if (columns[prop].Visible != false) {
                    header[ind] = new Array();
                    header[ind][0] = columns[prop].DisplayMemberName || columns[prop].Id;
                    header[ind][1] = columns[prop].Title || columns[prop].Text;
                    ind++;
                }
            }
        }

        for (var i = 0; i < header.length; i++) {
            this.AddHeaderMemberName(header[i][0]);
            this.AddHeader(header[i][1]);
        }

    };


    this.SetFromArray = function (dRow) {
        this.rows = new Array();
        this.rowCount = 0;
        for (var i = 0; i < dRow.length; i++) {
            this.AddNewRow();
            for (var j = 0; j < dRow[i].length; j++) {
                if (this.AllColumnsVisible || this.visibleColumns[j] == 'V') {
                    this.AddCell(dRow[i][j]);
                }
            };
        }

    };

    this.GetCellValue = function (row, column) {
        return this.rows[row - 1][column];
    };

    this.SetNthCellFormat = function (ind, value) {
        if (ind > 0) {
            for (var i = 0; i < ind; i++) {
                if (this.format[i] == null) {
                    this.format[i] = this.FormatType.String;
                }
            }
            this.format[ind - 1] = value;
        }
    };

    this.VisibleColumns = function (min, max) {
        if (min < 1) {
            min = 1;
        }
        this.AllColumnsVisible = false;
        if (max >= min) {
            for (var i = min; i <= max; i++) {
                this.visibleColumns[i - 1] = 'V';
            }
        }
    };
    // 
    this.CalculateColumnsSum = function (column) {

        this.ShowSumColumn = true;

        var total = 0;
        for (var i = 0; i < this.rowCount - 1; i++) {
            total += parseDecimal(this.rows[i][column - 1]);
        }

        this.rows[this.rowCount - 1][column - 1] = yeniStrReplace(total, ".", ",");
    };
}
//-------------------------------------------------------------------------------

function OpenExcel(obj) {
    obj.rowCount = 0;
    if (!obj.rows) {
        $Msg.myAlert("ExcelKolonTanimlariHatali");
        return false;
    }

    try {
        var oExcel = new JExcel();
        var oBook = oExcel.Workbooks.Add();
        var oSheet = oBook.Worksheets(1);
    } catch (err) {
        $Msg.myAlert(err && err.description ? err.description : err);
        return false;
    }

    if (obj.AddTopHeader) {
        obj.AddTopHeader(oSheet);
    }

    if (obj.AddSubHeader) {
        obj.AddSubHeader(oSheet);
    }

    if (obj.AddCellXY) {
        obj.AddCellXY(oSheet);
    }

    if (obj.CustomizeHeader) {
        obj.CustomizeHeader(oSheet);
    } else {
        if (obj.header != null) {
            //Headerı set edelim
            obj.rowCount++;
            for (var i = 0; i < obj.header.length; i++) {
                oSheet.Cells(obj.rowCount, i + 1).Value = obj.header[i];
            }
        }
    }
    obj.rowCount++;

    //Her bir satırı sırasıyla ekleyelim. 
    for (var i = 0; i < obj.rows.length; i++) {
        for (var j = 0; j < obj.rows[i].length; j++) {
            if (obj.rows[i][j] != undefined && obj.rows[i][j] != null) {
                if (obj.format != undefined && obj.format != null && obj.format.length >= j) {
                    oSheet.Cells(obj.rowCount, j + 1).NumberFormat = obj.format[j];
                    if (obj.format[j] == obj.FormatType.Number) {
                        oSheet.Cells(obj.rowCount, j + 1).Value = parseDecimal(obj.rows[i][j]);
                    }
                    else if (obj.format[j] == obj.FormatType.Date) {
                        oSheet.Cells(obj.rowCount, j + 1).Value = obj.rows[i][j].replace(/\//g, '.');
                    }
                    else {
                        oSheet.Cells(obj.rowCount, j + 1).Value = obj.rows[i][j];
                    }
                } else {
                    oSheet.Cells(obj.rowCount, j + 1).NumberFormat = obj.FormatType.String;
                    oSheet.Cells(obj.rowCount, j + 1).Value = obj.rows[i][j];
                }
            } else {
                oSheet.Cells(obj.rowCount, j + 1).NumberFormat = obj.FormatType.String;
                oSheet.Cells(obj.rowCount, j + 1).Value = '';
            }
        }
        obj.rowCount++;
    }

    if (obj.header != null) {
        for (var i = 0; i < obj.header.length; i++) {
            oSheet.Cells(1, i + 1).font.bold = true;
            oSheet.Cells(1, i + 1).Interior.ColorIndex = 40;
        }
    }

    if (obj.ShowSumColumn) {
        var rc = (obj.rowCount - 1).toString();
        var rowNumber = obj.rows.length + 1;
        for (var i = 0; i < obj.header.length; i++) {
            oSheet.Cells(rowNumber, i + 1).Interior.ColorIndex = 40;
            oSheet.Cells(rowNumber, i + 1).font.bold = true;

            if (obj.format[i] == obj.FormatType.Number) {
                oSheet.Cells(rc, i + 1).NumberFormat = obj.format[i];
            }

        }
    }

    oSheet.Columns.AutoFit();
    oExcel.Visible = true;
    oExcel.UserControl = true;
}
//-------------------------------------------------------------------------
// Excele çıkacak kayıtlar için öncelikle bu functiondan obje oluşturulacak 
function ExcelMaster() {

    this.FormatType = {
        Number: "0.00",
        Date: "dd.mm.yyyy",
        Decimal: "0",
        Currency: isIE() ? "$" : "#,##0.00 ₺;[Red]-#,##0.00 ₺",
        String: "@"
    };

    this.oExcel = null;
    this.oBook = null;
    this.oSheet = null;

    this.Initialize = function () {
        this.header = new Array();
        this.headerMemberName = new Array();
        this.format = new Array();
        this.rows = new Array();
        this.visibleColumns = new Array();
        this.rowCount = 1;
        this.colCount = 1;
        this.currentFormat = this.FormatType.String;
        this.AllColumnsVisible = true;
        this.ShowSumColumn = false;
        this.addHeaderOnlyText = true;
    }
    this.Initialize();
    // Excel açılırken bir defaya mahsus çağrılır ve sheetlerin fazlalarını siler.
    this.Open = function () {
        if (this.oExcel != null) {
            return false;
        }
        try {
            this.oExcel = new JExcel();
            this.oBook = this.oExcel.Workbooks.Add();
            this.oSheet = this.oBook.Worksheets(1);
        } catch (err) {
            myAlert(err && err.description ? err.description : err);
            return false;
        }
    }

    //excel dosyasının ismini verir
    this.SetWorkbookName = function (name) {
        if (!fora.Inue(name)) {
            this.oExcel.Caption = name;
        }
    }

    // Yeni sheet eklemek için kullanılır
    this.OpenNewSheet = function (sheetName) {
        for (var i = 0, len = this.header.length; i < len; i++) {
            this.oSheet.Cells(1, i + 1).Value = this.header[i];
            this.oSheet.Cells(1, i + 1).Font.Bold = true;
            this.oSheet.Cells(1, i + 1).Interior.ColorIndex = 40;
        }
        this.Open();
        this.Initialize();
        this.oSheet = this.oBook.Worksheets.Add();
        this.SetActiveSheetName(sheetName);
    }

    // Excel Sheete isim ataması yapar
    this.SetActiveSheetName = function (sheetName) {
        if (sheetName != null || sheetName != undefined) {
            this.oSheet.Name = sheetName;
        }
    }

    this.Show = function () {
        for (var i = 0, len = this.header.length; i < len; i++) {
            this.oSheet.Cells(1, i + 1).Value = this.header[i];
            this.oSheet.Cells(1, i + 1).Font.Bold = true;
            this.oSheet.Cells(1, i + 1).Interior.ColorIndex = 40;
        }

        for (var i = 1; i <= this.oBook.Worksheets.Count; i++) {
            this.oBook.Worksheets(i).Columns.AutoFit();
        }

        this.oSheet.Columns.AutoFit();
        //Birden fazla sheet varsa ilkini aktif edelim    
        this.oBook.Worksheets(1).Activate();
        this.oExcel.Visible = true;
        this.oExcel.UserControl = true;
    }

    this.GetCurrentRow = function () {
        return this.rowCount;
    }
    this.GetCurrentColumn = function () {
        return this.colCount;
    }
    this.SetCurrentRow = function (ind) {
        this.rowCount = ind;
    }
    this.SetCurrentColumn = function (ind) {
        this.colCount = ind;
    }

    this.AddHeader = function (value) {
        this.header[this.header.length] = value;
    };

    this.AddHeaderMemberName = function (value) {
        this.headerMemberName[this.headerMemberName.length] = value.replace(/__/g, '.');
    };

    this.AddFormat = function (value) {
        this.format[this.format.length] = value;
    };

    this.AddNewRow = function () {
        //this.rows[this.rowCount] = new Array();
        this.rowCount++;
        this.colCount = 1;
    };
    this.bold = false;

    this.AddCells = function (value, count, format, colorInd, bold) {
        for (var i = 0; i < count; i++) {
            this.AddCell(value, format, colorInd, bold);
        }
    }

    this.AddCell = function (value, format, colorInd, bold) {
        if (this.rowCount == 0 || (this.header.length > 0 && this.rowCount == 1)) {
            this.AddNewRow();
        }
        // seçili hücrenin formatını belirler
        if (format == undefined || format == null) {
            this.oSheet.Cells(this.rowCount, this.colCount).NumberFormat = this.FormatType.String;
        } else {
            this.oSheet.Cells(this.rowCount, this.colCount).NumberFormat = format;
        }

        // seçili hücrenin rengini atar.
        if (colorInd != undefined && colorInd != null) {
            this.oSheet.Cells(this.rowCount, this.colCount).Interior.ColorIndex = colorInd;
        }
        // hücrenin fontunu kalın yapar
        if (bold == undefined || bold == null) {
            this.oSheet.Cells(this.rowCount, this.colCount).font.bold = false;
        } else {
            this.oSheet.Cells(this.rowCount, this.colCount).font.bold = bold;
        }
        if (format == this.FormatType.Number || format == this.FormatType.Decimal) {
            this.oSheet.Cells(this.rowCount, this.colCount++).Value = parseDecimal(value);
        } else {
            this.oSheet.Cells(this.rowCount, this.colCount++).Value = value;
        }
    };

    this.SetHeader = function (arr) {
        this.header = arr;
    };

    this.SetHeaderMemberName = function (arr) {
        this.headerMemberName = arr;
    };

    //Excele çıkacak kayıtlar için Datagriddeki başlıkları alıyoruz.
    this.AddHeaderFromDataGrid = function (columns, colorInd, bold) {

        // seçili hücrenin rengini atar. http://www.mvps.org/dmcritchie/excel/colors.htm
        if (colorInd == undefined || colorInd == null) {
            colorInd = 40;
        }
        // hücrenin fontunu kalın yapar
        if (bold == undefined && bold == null) {
            bold = true;
        }
        var header = new Array();
        var ind = 0;
        for (var prop in columns) {
            if (columns.hasOwnProperty(prop)) {
                if (columns[prop].Visible != false && (!this.addHeaderOnlyText || (this.addHeaderOnlyText && columns[prop].Type != "Control"))) {
                    header[ind] = new Array();
                    header[ind][0] = columns[prop].DisplayMemberName || columns[prop].Id;
                    header[ind][1] = columns[prop].Title || columns[prop].Text;
                    header[ind][2] = columns[prop].Type || columns[prop].ForaType;
                    this.AddCell(header[ind][1], null, colorInd, bold);
                    ind++;
                }
            }
        }
        this.header = new Array();
        this.headerMemberName = new Array();
        this.format = new Array();
        for (var i = 0; i < header.length; i++) {
            this.AddHeaderMemberName(header[i][0]);
            this.AddHeader(header[i][1]);
            this.AddFormat(header[i][2]);
        }
    };


    //Excele çıkacak kayıtlar için Datagriddeki başlıkları alıyoruz.
    this.AddHeaderFromArray = function (arr, colorInd, bold) {

        // seçili hücrenin rengini atar. http://www.mvps.org/dmcritchie/excel/colors.htm
        if (colorInd == undefined || colorInd == null) {
            colorInd = 40;
        }
        // hücrenin fontunu kalın yapar
        if (bold == undefined && bold == null) {
            bold = true;
        }
        var header = new Array();
        var ind = 0;
        for (var prop in arr) {
            if (arr.hasOwnProperty(prop)) {
                header[ind] = new Array();
                header[ind][0] = arr[prop][0];
                header[ind][1] = arr[prop][1];
                header[ind][2] = arr[prop][2];
                this.AddCell(arr[prop][1], null, colorInd, bold);
                ind++;

            }
        }
        this.header = new Array();
        this.headerMemberName = new Array();
        this.format = new Array();
        for (var i = 0; i < header.length; i++) {
            this.AddHeaderMemberName(header[i][0]);
            this.AddHeader(header[i][1]);
            this.AddFormat(header[i][2]);
        }
    };

    this.AddDataGridRows = function (data, colorInd, bold, selectedRowIndex, selectedColumnIndex) {
        if (data == null || data.length == 0) {
            return;
        }
        if (colorInd == undefined) {
            colorInd = null;
        }
        if (bold == undefined) {
            bold = false;
        }
        if (selectedRowIndex == undefined || selectedRowIndex == null) {
            selectedRowIndex = -1;
        }
        if (selectedColumnIndex == undefined || selectedColumnIndex == null) {
            selectedColumnIndex = -1;
        }
        var orgColorInd = colorInd;
        var dgStartColInd = this.GetCurrentColumn();
        for (var rowCount = 0; rowCount < data.length; rowCount++) {
            this.SetCurrentColumn(dgStartColInd);
            if (selectedRowIndex > -1 && rowCount == selectedRowIndex && selectedColumnIndex == -1) {
                colorInd = 43;
            }
            for (var colCount = 0; colCount < this.header.length; colCount++) {
                if ((rowCount == selectedRowIndex && selectedColumnIndex == -1) || (rowCount == selectedRowIndex && colCount == selectedColumnIndex)) {
                    colorInd = 43;
                } else {
                    colorInd = orgColorInd;
                }
                // DisplayMemberName="BranchCode|||Number|||Suffix" şeklindeki kolonlar için özel kontrol eklendi.
                var colText = this.headerMemberName[colCount];
                if (colText.indexOf('|||') > -1) {
                    var arr = colText.split('|||');
                    colText = '';
                    for (var i = 0; i < arr.length; i++) {
                        if (i == 0) {
                            colText = data[rowCount][arr[i]];
                        } else {
                            colText += '-' + data[rowCount][arr[i]];
                        }
                    }
                } else {
                    colText = data[rowCount][this.headerMemberName[colCount]];
                }
                if (this.format[colCount] == "Number") {
                    this.AddCell(colText, this.FormatType.Number, colorInd, bold);
                } else if (this.format[colCount] == "Decimal") {
                    this.AddCell(colText, this.FormatType.Decimal, colorInd, bold);
                } else if (this.format[colCount] == "DateTime") {
                    colText = colText.replace(/\//g, '.');
                    this.AddCell(colText, this.FormatType.Date, colorInd, bold);
                } else {
                    this.AddCell(colText, null, colorInd, bold);
                }
                if (selectedColumnIndex == colCount) {
                    colorInd = orgColorInd;
                }
            }
            this.AddNewRow();
            colorInd = orgColorInd;
        }
    };

    this.SetHeaderFormat = function (ind, value) {
        if (this.format == null || this.format.length < ind) {
            return false;
        }
        this.format[ind] = value;
    }

    this.SetFromArray = function (dRow) {
        this.rows = new Array();
        //this.rowCount = 0;
        for (var i = 0; i < dRow.length; i++) {
            for (var j = 0; j < dRow[i].length; j++) {
                if (this.AllColumnsVisible || this.visibleColumns[j] == 'V') {
                    this.AddCell(dRow[i][j]);
                }
            };
            this.AddNewRow();
        }
    };

    this.GetCellValue = function (row, column) {
        return this.rows[row - 1][column];
    };

    this.SetNthCellFormat = function (ind, value) {
        if (ind > 0) {
            for (var i = 0; i < ind; i++) {
                if (this.format[i] == null) {
                    this.format[i] = this.FormatType.String;
                }
            }
            this.format[ind - 1] = value;
        }
    };

    this.VisibleColumns = function (min, max) {
        if (min < 1) {
            min = 1;
        }
        this.AllColumnsVisible = false;
        if (max >= min) {
            for (var i = min; i <= max; i++) {
                this.visibleColumns[i - 1] = 'V';
            }
        }
    };
}
