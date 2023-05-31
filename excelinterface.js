//Bu dosyadaki sınıflar IE modunu bozduğundan babeljs ile compile edilip ordan elde edilen çıktı kullanılmalı. IE hayatımızdan çıktığında doğrudan bu dosyayı faaliyete alacağız

class ExcelIO {
    /**
     * Url'i verilen dosyayı açar
     * @param {any} url dosya linki
     * @param {any} callback dosya açıldıktan sonra çağrılacak fonksiyon. Parametre olarak Excel nesnesini geçer
     */
    static Open(url, callback) {
        if (!callback || typeof callback != "function") {
            WriteError("Open için callback fonksiyonu bulunamadı!");
            return;
        }
        var excelApp = new JExcel();
        if (isIE()) {
            excelApp.Workbooks.Open(url);
            callback(excelApp);
        }
        else {
            excelApp.Open(url, function () {
                callback(excelApp);
            })
        }
        return true;
    }

    /**
     * Yerel sistemde yer alan dosyayı okumak için kullanılır.
     * @param {any} callback dosya açıldıktan sonra çağrılacak fonksiyon. Parametre olarak Excel nesnesini geçer
     */
    static ReadFile(callback) {
        if (!callback || typeof callback != "function") {
            WriteError("ReadFile için callback fonksiyonu bulunamadı!");
            return;
        }
        var excelApp = new JExcel();
        if (isIE()) {
            excelApp.FindFile();
            callback(excelApp);
        }
        else {
            excelApp.FindFile(function (workbook) {
                callback(excelApp);
            });
        }
        return true;
    }

    /**
     * Aldığı workbook objesini excel olarak export eder
     * @param {ExcelBook} wb 
     */
    static Export(wb) {
        if (wb.Worksheets.length == 0) {
            WriteError("Export edilecek sayfa bulunamadı");
            return false;
        }
        var oExcel, oBook;
        try {
            oExcel = new JExcel();
            oBook = oExcel.ActiveWorkbook || oExcel.Workbooks.Add();
            //oExcel.SheetsInNewWorkbook = wb.Worksheets.length;         
        } catch (err) {
            WriteError(err && err.description ? err.description : err);
            return false;
        }

        oExcel.Caption = wb.Name;
        wb.Worksheets.forEach(function (v, i) {
            var oSheet = oBook.Worksheets.Add();
            var sh = wb.Worksheets[i];
            oSheet.Name = sh.Name;

            if (sh.AddTopHeader) {
                sh.AddTopHeader(oSheet);
            }

            if (sh.AddSubHeader) {
                sh.AddSubHeader(oSheet);
            }

            if (sh.AddCellXY) {
                sh.AddCellXY(oSheet);
            }

            var hBreak = 0;
            if (sh.CustomizeHeader) {
                sh.CustomizeHeader(oSheet);
                if (oSheet.UsedRange.UsedRange.Rows.Count > 0)
                    hBreak = 1;
            }
            else {
                if (sh.Headers != null && sh.Headers.length > 0) {
                    hBreak = 1;
                    for (var h = 0; h < sh.Headers.length; h++) {
                        oSheet.Cells(1, h + 1).Value = sh.Headers[h].Title;
                        oSheet.Cells(1, h + 1).Font.Bold = sh.Headers[h].Font && sh.Headers[h].Font.Bold || false;
                        oSheet.Cells(1, h + 1).Interior.ColorIndex = sh.Headers[h].Background;
                    }
                }
            }

            //hücreleri birleştirelim
            for (var m = 0; m < sh.Merges.length; m++) {
                var merge = sh.Merges[m];
                oSheet.mergeCells(merge.FromRow, merge.FromColumn, merge.ToRow, merge.ToColumn);
                oSheet.getCell(merge.FromRow, merge.FromColumn).alignment = { horizontal: 'center', vertical: 'middle' };
            }

            //Her bir satırı sırasıyla ekleyelim. 
            var r = 0;
            for (r = 0; r < sh.Cells.length; r++) {
                if (!sh.Cells[r])
                    continue;
                for (var c = 0; c < sh.Cells[r].length; c++) {
                    var cell = sh.Cells[r][c];
                    if (cell instanceof ExcelCell) {
                        var oCell = oSheet.Cells(r + 1 + hBreak, c + 1);
                        if (!fora.Inu(cell)) {
                            if (!fora.Inu(cell.Format)) {
                                oCell.NumberFormat = cell.Format;
                                if (cell.Format == ExcelCellFormatTypes.Number) {
                                    oCell.Value = parseDecimal(cell.Value);
                                } else {
                                    oCell.Value = cell.Value;
                                }
                            } else {
                                oCell.NumberFormat = ExcelCellFormatTypes.String;
                                oCell.Value = cell.Value;
                            }
                            if (cell.HasBackground)
                                oCell.Interior.ColorIndex = cell.Background;
                            //hücreye ait font var mı
                            oCell.Font.Bold = cell.Font && cell.Font.Bold || false;
                        } else {
                            oCell.NumberFormat = ExcelCellFormatTypes.String;
                            oCell.Value = '';
                        }
                    }

                }
            }

            for (var h = 0; h < sh.Headers.length; h++) {

                if (sh.Headers[h].ShowSumInfo) {
                    oSheet.Cells(r + hBreak + 1, h + 1).Value = { formula: `SUM(${ExcelHelpers.GetColumnLetter(h + 1)}${hBreak + 1}:${ExcelHelpers.GetColumnLetter(h + 1)}${r + hBreak})`, date1904: false };
                    oSheet.Cells(r + hBreak + 1, h + 1).Font.Bold = true;
                    oSheet.Cells(r + hBreak + 1, h + 1).Interior.ColorIndex = 40;
                    oSheet.Cells(r + hBreak + 1, h + 1).NumberFormat = sh.Headers[h].Format;
                }
            }

            oSheet.Columns.AutoFit();
        })

        oExcel.Visible = true;
        return true;
    }
}


class ExcelCellFormatTypes {
    static get Number() { return "0.00" };
    static get Date() { return "dd.mm.yyyy" };
    static get Decimal() { return "0" };
    static get Currency() { return "#,##0.00 ₺;[Red]-#,##0.00 ₺" };
    static get String() { return "" };
}

class ExcelBook {
    /**
     * Yeni ExcelBook oluşturur
     * @param {string} name Oluşturulacak kitabın adı
     */
    constructor(name) {
        /** @type {string} */
        this.Name = name || "Workbook";

        /** @type {Array<ExcelSheet>} */
        this.Worksheets = [];
    }

    /**
     * Yeni ExcelBook oluşturur
     * @param {string} name Oluşturulacak kitabın adı
     */
    static Create(name) {
        return new ExcelBook(name);
    }

    /**
     * Kitaba yeni sayfa ekler
     * @param {string} sheetName Sayfa adı
     */
    AddSheet(sheetName) {
        let sheet = new ExcelSheet(this.Worksheets.length + 1, sheetName)
        this.Worksheets.push(sheet);
        return sheet;
    }

    /**
     * Verilen kritere uyan sayfayı getirir
     * @param {number | string} criteria
     * @returns {ExcelSheet}
     */
    GetSheet(criteria) {
        if (!criteria)
            WriteError("Not found criteria to get sheet!");
        let [result] = this.Worksheets.filter(x => x.Id == criteria)
        if (!result)
            [result] = this.Worksheets.filter(x => x.Name == criteria)
        return result;
    }

    /**
     * Sayfayı yeniden isimlendirir
     * @param {string} name
     */
    Rename(name) {
        this.Name = name;
        return this;
    }

    /**
     * Verilen JSON veriyi kitaba dönüştürür.
     * @param {JSON} data
     */
    static JsonToWorkbook(data) {


        let checkAllItemTypeIsSame = function (data) {
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
        /**
         * 
         * @param {any} data
         * @param {ExcelSheet} sheet
         */
        let dataToExcel = function (data, sheet) {
            var cols = [];
            for (var i = 0; i < data.length; i++) {
                for (var k in data[i]) {
                    if (cols.indexOf(k) == -1) {
                        cols.push(k);
                    }
                }
            }

            cols.forEach(function (id) {
                sheet.AddHeader(id, fora.Translate(id), new ExcelStyle(undefined, 43, true));
            });
            data.forEach(function (item, index) {
                var cells = [];
                cols.forEach(function (key) {
                    cells.push(new ExcelCell(item[key]));
                });
                if (index == 0)
                    sheet.AddCells(cells);
                else
                    sheet.AddNewRow(cells)
            })
        }
        if (fora.Inue(data)) {
            WriteError("Excel'e aktarılacak veri bulunamadı.");
        }
        else if (!Array.isArray(data)) {
            WriteError("Beklenmedik veri formatı. Format [...] ya da [[...],[...]] şeklinde olmalıdır.");
        }
        else if (data.length <= 0) {
            WriteError("Excel'e aktarılacak veri bulunamadı.");
        }
        else if (!checkAllItemTypeIsSame(data)) {
            WriteError("Beklenmedik veri formatı. Format [...] ya da [[...],[...]] şeklinde olmalıdır.");
        }

        var book = new ExcelBook();

        if (Array.isArray(data[0])) {
            for (var i = 0; i < data.length; i++) {
                if (!checkAllItemTypeIsSame(data[i])) {
                    WriteError("Beklenmedik veri formatı. Format [...] ya da [[...],[...]] şeklinde olmalıdır.");
                }
                let sheet = book.AddSheet();
                dataToExcel(data[i], sheet);
            }
        }
        else {
            let sheet = book.AddSheet();
            dataToExcel(data, sheet);
        }
        return book;
    }
}

class ExcelKnownColors {
    static get Black() { return 1; }
    static get White() { return 2; }
    static get Red() { return 3; }
    static get Green() { return 4; }
    static get Blue() { return 5; }
    static get Yellow() { return 6; }
    static get Magenta() { return 7; }
    static get Cyan() { return 8; }
    static get ColorOf800000() { return 9; }
    static get ColorOf008000() { return 10; }
    static get ColorOf000080() { return 11; }
    static get ColorOf808000() { return 12; }
    static get ColorOf800080() { return 13; }
    static get ColorOf008080() { return 14; }
    static get ColorOfC0C0C0() { return 15; }
    static get ColorOf808080() { return 16; }
    static get ColorOf9999FF() { return 17; }
    static get ColorOf993366() { return 18; }
    static get ColorOfFFFFCC() { return 19; }
    static get ColorOfCCFFFF() { return 20; }
    static get ColorOf660066() { return 21; }
    static get ColorOfFF8080() { return 22; }
    static get ColorOf0066CC() { return 23; }
    static get ColorOfCCCCFF() { return 24; }
    static get ColorOf000080() { return 25; }
    static get ColorOfFF00FF() { return 26; }
    static get ColorOfFFFF00() { return 27; }
    static get ColorOf00FFFF() { return 28; }
    static get ColorOf800080() { return 29; }
    static get ColorOf800000() { return 30; }
    static get ColorOf008080() { return 31; }
    static get ColorOf0000FF() { return 32; }
    static get ColorOf00CCFF() { return 33; }
    static get ColorOfCCFFFF() { return 34; }
    static get ColorOfCCFFCC() { return 35; }
    static get ColorOfFFFF99() { return 36; }
    static get ColorOf99CCFF() { return 37; }
    static get ColorOfFF99CC() { return 38; }
    static get ColorOfCC99FF() { return 39; }
    static get ColorOfFFCC99() { return 40; }
    static get ColorOf3366FF() { return 41; }
    static get ColorOf33CCCC() { return 42; }
    static get ColorOf99CC00() { return 43; }
    static get ColorOfFFCC00() { return 44; }
    static get ColorOfFF9900() { return 45; }
    static get ColorOfFF6600() { return 46; }
    static get ColorOf666699() { return 47; }
    static get ColorOf969696() { return 48; }
    static get ColorOf003366() { return 49; }
    static get ColorOf339966() { return 50; }
    static get ColorOf003300() { return 51; }
    static get ColorOf333300() { return 52; }
    static get ColorOf993300() { return 53; }
    static get ColorOf993366() { return 54; }
    static get ColorOf333399() { return 55; }
    static get ColorOf333333() { return 56; }
}

class ExcelStyle {
    /**
     * Yeni stil nesnesi oluşturur.
     * @param {string} format
     * @param {ExcelKnownColors | number} colorIndex
     * @param {boolean} isBold
     */
    constructor(format, colorIndex, isBold) {

        this.Format = format || ExcelCellFormatTypes.String;
        this.Background = colorIndex;
        this.Font = { Bold: !!isBold };
    }

    /**
     * Yeni stil nesnesi oluşturur.
     * @param {string} format
     * @param {ExcelKnownColors | number} colorIndex
     * @param {boolean} isBold
     */
    static Create(format, colorIndex, isBold) {
        return new ExcelStyle(format, colorIndex, isBold);
    }

    get HasBackground() {
        return this.Background !== undefined
    }
}


class ExcelHeader extends ExcelStyle {
    /**
     * Yeni bir ExcelHeader nesnesi oluşturur.
     * @param {string} id Başlık kimliği
     * @param {string} title Başlık değeri
     * @param {ExcelStyle} style Başlık hücresine ait stiller
     */
    constructor(id, title, style = {}) {
        super(style.Format, style.Background || ExcelKnownColors.ColorOfFFCC99, (style.Font || { Bold: true }).Bold);
        /** @type {string} */
        this.Id = id;
        /** @type {string} */
        this.Title = title || id;
        /** @type {boolean} */
        this.ShowSumInfo = false;

    }

    ShowSum() {
        this.ShowSumInfo = true;
        return this;
    }

    /**
     * Yeni bir ExcelHeader nesnesi oluşturur.
     * @param {string} id Başlık kimliği
     * @param {string} title Başlık değeri
     * @param {ExcelStyle} style Başlık hücresine ait stiller
     */
    static Create(id, title, style = {}) {
        return new ExcelHeader(id, title, style);
    }
}

class ExcelCell extends ExcelStyle {
    /**
     * Yeni bir ExcelCell nesnesi oluşturur.
     * @param {any} value
     * @param {ExcelStyle} style
     */
    constructor(value, style = {}) {
        super(style.Format, style.Background, style.Font && style.Font.Bold || false);
        this.Value = value;
    }

    /**
      * Yeni bir ExcelCell nesnesi oluşturur.
      * @param {any} value
      * @param {ExcelStyle} style
      */
    static Create(value, style = {}) {
        return new ExcelCell(value, style);
    }
}

class ExcelHelpers {
    /**
     * Verilen tipe uygun excel formatını döner
     * @param {string} type
     */
    static GetExcelFormatOfType(type) {
        switch (type) {
            case "Date":
            case "DateTime":
                return ExcelCellFormatTypes.Date;
            case "Number":
                return ExcelCellFormatTypes.Number;
            case "Decimal":
                return ExcelCellFormatTypes.Decimal;
            case "Money":
                return ExcelCellFormatTypes.Currency;
            default:
                return ExcelCellFormatTypes.String;
        }
    }

    /**
     * Verilen değerin, belirtilen aralıkta olup olmadığını kontrol eder.
     * @param {number} v Aranan sayı
     * @param {number} r1 Aralık başlangıcı (>=)
     * @param {number} r2 Aralık bitişi (<=)
     */
    static IsInRange(v, r1, r2) {
        var max = Math.max(r1, r2);
        var min = Math.min(r1, r2);
        return min <= v && v <= max;
    }

    /**
     * Verilen sütun numarasına göre sütun etiketi oluşturur
     * @param {number} columnNumber 1 ve 1 den büyük olmalı
     */
    static GetColumnLetter(columnNumber) {
        var n = columnNumber - 1;
        var ordA = 'A'.charCodeAt(0);
        var ordZ = 'Z'.charCodeAt(0);
        var len = ordZ - ordA + 1;

        var s = "";
        while (n >= 0) {
            s = String.fromCharCode(n % len + ordA) + s;
            n = Math.floor(n / len) - 1;
        }
        return s;
    }
}

class MergeInfo {
    /**
     * Birleştirilmiş hücrelerin bilgisini tutar.
     * @param {ExcelCell} actualCell Esas veri hücresi
     * @param {number} fromRow Birleşim başlangıç satırı
     * @param {number} fromColumn Birleşim başlangıç sütunu
     * @param {number} toRow Birleşim bitiş satırı
     * @param {number} toColumn Birleşim bitiş sütunu
     */
    constructor(actualCell, fromRow, fromColumn, toRow, toColumn) {
        /** @type {ExcelCell} */
        this.ActualCell = actualCell;
        /** @type {number} */
        this.FromRow = fromRow;
        /** @type {number} */
        this.FromColumn = fromColumn;
        /** @type {number} */
        this.ToRow = toRow;
        /** @type {number} */
        this.ToColumn = toColumn;
    }
}

class ExcelSheet {

    /**
     * 
     * @param {string} id Sayfa kimliği
     * @param {string} name Sayfa Adı
     */
    constructor(id, name) {
        /** @type {string} */
        this.Id = id;
        /** @type {string} */
        this.Name = name || `Sheet ${id}`;
        /** @type {Array<ExcelHeader>} */
        this.Headers = [];
        /** @type {Array<ExcelCell>} */
        this.Cells = [];
        /** @type {Array<MergeInfo>} */
        this.Merges = [];
    }

    /**
     *
     * @param {string} id Sayfa kimliği
     * @param {string} name Sayfa Adı
     */
    static Create(id, name) {
        return new ExcelSheet(id, name)
    }

    /**
     * Sayfayı yeniden isimlendirir.
     * @param {string} name
     */
    Rename(name) {
        this.Name = name;
        return this;
    }

    /**
     * Verilen değerlerle ilk satırda yer alacak başlık hücrelerini ekler
     * @param {string} id
     * @param {string} title
     * @param {ExcelStyle} style
     */
    AddHeader(id, title, style = {}) {
        var header = new ExcelHeader(id, title, style);
        this.Headers.push(header);
        return header;
    }

    #r = 0;
    #c = 0;

    /**
     * Hücreyi verilen değerle doldurur
     * @param {any} value Hücreye yazılacak değer
     * @param {number} row Güncellenecek hücrenin satır numarası
     * @param {number} column Güncellenecek hücrenin sütun numarası
     * @param {ExcelStyle} style Hücre stili
     */
    AddCell(value, row, column, style = {}) {
        if (fora.Inu(row) || fora.Inu(column)) {
            row = this.#r;
            column = this.#c;
        }
        else {
            row--;
            column--;
        }

        var r = this.Cells[row];
        if (!r)
            this.Cells[row] = [];
        var x = this.Cells[row][column];
        if (x instanceof ExcelCell) {
            x.Value = value;
            x.Background = style.Background;
            x.Font = style.Font || { Bold: false };
            x.Format = style.Format;
        }
        else {
            this.Cells[row][column] = new ExcelCell(value, style);
        }

        column++;

        if (row > this.#r) {
            this.#r = row;
            this.#c = column;
        }
        if (column > this.#c)
            this.#c = column;
    }

    /**
     * Yeni boş bir satır ekler, values değeri ExcelCell tipinde dolu şekilde geldiyse yeni eklenen satıra bu veriyi ekler
     * @param {...ExcelCell} values
     */
    AddNewRow(...values) {
        let cellValues;
        if (arguments.length === 1 && Array.isArray(arguments[0]))
            cellValues = arguments[0];
        else
            cellValues = values;

        this.Cells[++this.#r] = [];
        this.#c = 0;
        if (cellValues instanceof Array && cellValues.length > 0) {
            if (cellValues.every(v => v instanceof ExcelCell)) {
                cellValues.forEach(v => {
                    this.Cells[this.#r][this.#c++] = v;
                })
            }
        }
    }

    /**
     * 
     * @param {...ExcelCell} values
     */
    AddCells(...values) {
        let cellValues;
        if (arguments.length === 1 && Array.isArray(arguments[0]))
            cellValues = arguments[0];
        else
            cellValues = values;
        if (cellValues instanceof Array && cellValues.length > 0) {
            if (cellValues.every(v => v instanceof ExcelCell)) {
                cellValues.forEach(v => {
                    this.AddCell(v.Value, undefined, undefined, new ExcelStyle(v.Format, v.Background, v.Font && v.Font.Bold || false));
                })
            }
        }
    }

    /**
     * Verilen aralıktaki hücreleri birleştirir. 
     * @param {any} fromRow
     * @param {any} fromCol
     * @param {any} toRow
     * @param {any} toCol
     */
    MergeCells(fromRow, fromCol, toRow, toCol) {
        var self = this;
        //3, 2, 5, 3

        function hasConflict() {
            var result = self.Merges.find(x => {

                var rowConflict = ExcelHelpers.IsInRange(x.FromRow, fromRow, toRow) || ExcelHelpers.IsInRange(x.ToRow, fromRow, toRow);
                var colConflict = ExcelHelpers.IsInRange(x.FromColumn, fromCol, toCol) || ExcelHelpers.IsInRange(x.ToColumn, fromCol, toCol);
                if (rowConflict && colConflict)
                    return true;
            });
            return !!result;
        }

        if (hasConflict())
            WriteError(`Birleştirilmek istenen alanda daha önceden birleştirilmiş kayıtlar mevcut. Alanı değiştirip tekrar deneyin. Alanınızın başlangıç satırı ${fromRow}, başlangıç sütunu ${fromCol}, bitiş satırı ${toRow}, bitiş sütunu ${toCol}`);
        var firstCell = null;
        for (var r = fromRow - 1; r < toRow; r++) {
            for (var c = fromCol - 1; c < toCol; c++) {
                if (this.Cells[r] && this.Cells[r][c]) {
                    firstCell = this.Cells[r][c];
                    r = toRow;
                    break;
                }
            }
        }
        if (!firstCell)
            firstCell = new ExcelCell("");
        this.Merges.push(new MergeInfo(firstCell, fromRow, fromCol, toRow, toCol));

        for (var r = fromRow - 1; r < toRow; r++) {
            for (var c = fromCol - 1; c < toCol; c++) {
                if (!this.Cells[r]) {
                    this.Cells[r] = [];
                }
                this.Cells[r][c] = firstCell;
            }
        }
    }


    UnMergeCell(row, col) {
        var mergeInfos = this.Merges.filter(x => ExcelHelpers.IsInRange(row, x.FromRow, x.ToRow) && ExcelHelpers.IsInRange(col, x.FromColumn, x.ToColumn));
        mergeInfos.forEach(x => {
            for (var i = x.FromRow - 1; i < x.ToRow; i++) {
                var row = this.Cells[i];
                if (row instanceof Array) {
                    row.splice(x.FromColumn - 1, x.ToColumn - x.FromColumn + 1)
                }
            }
            this.Cells[x.FromRow - 1][x.FromColumn - 1] = x.ActualCell;
            var index = this.Merges.indexOf(x);
            this.Merges.splice(index, 1);
        })

    }

    /**
     * Verilen adres aralığındaki hücrelerin stillerini günceller
     * @param {number} fromRow Başlangıç satırı
     * @param {number} fromCol Başlangıç sütunu
     * @param {number} toRow Bitiş satırı
     * @param {number} toCol Bitiş sütunu
     * @param {ExcelStyle} style Hücre stili
     */
    SetRangeStyle(fromRow, fromCol, toRow, toCol, style) {
        if (!(style instanceof ExcelStyle))
            WriteError("style parametresi ExcelStyle tipinde olmalıdır.")
        for (let r = fromRow - 1; r < toRow; r++) {
            var rValues = this.Cells[r];
            if (rValues) {
                for (let index = (r == fromRow - 1 ? fromCol - 1 : 0); index < (r == toRow - 1 ? toCol : rValues.length); index++) {
                    const cell = rValues[index];
                    if (cell instanceof ExcelCell) {

                        if (cell.Font !== undefined)
                            cell.Font = style.Font;
                        if (style.HasBackground)
                            cell.Background = style.Background;
                        if (style.Format !== undefined)
                            cell.Format = style.Format;
                    }
                }
            }
        }
    }

    /**
     * Verilen adresteki hücrenin stillerini günceller
     * @param {number} row Hücrenin yer aldığı satır
     * @param {number} column Hücrenin yer aldığı sütun
     * @param {ExcelStyle} style Hücre stili
     */
    SetCellStyle(row, column, style) {
        this.SetRangeStyle(row, column, row, column, style);
    }

    /**
     * Belirtilen hücreyi döner
     * @param {number} row Hücrenin yer aldığı satır numarası
     * @param {number} column Hücrenin yer aldığı satır numarası
     * @returns {ExcelCell}
     */
    GetCell(row, column) {
        return this.Cells[row - 1] && this.Cells[row - 1][column - 1];
    };

    /**
     * Belirtilen hücrenin değerini döner
     * @param {number} row Hücrenin yer aldığı satır numarası
     * @param {number} column Hücrenin yer aldığı satır numarası
     */
    GetCellValue(row, column) {
        return this.Cells[row - 1] && this.Cells[row - 1][column - 1] && this.Cells[row - 1][column - 1].Value;
    };

    /**
     * 
     * @param {any} gridComponent
     * @param {ExcelKnownColors} colorIndex
     * @param {boolean} isBold
     * @param {string} format
     */
    SetHeadersFromDataGrid(gridComponent, colorIndex, isBold = true) {
        if (!gridComponent)
            WriteError("gridComponent boş olmamalı ve Grid tipinde olmalı");
        if (gridComponent.Columns && gridComponent.Columns instanceof Array) {
            //Interframe Grid
            let columns = gridComponent.Columns;
            this.Headers = [];
            for (var prop in columns) {
                if (columns.hasOwnProperty(prop)) {
                    if (columns[prop].Visible != false && (!this.addHeaderOnlyText || (this.addHeaderOnlyText && columns[prop].Type != "Control"))) {
                        this.AddHeader(columns[prop].DisplayMemberName, columns[prop].Title, new ExcelStyle(ExcelHelpers.GetExcelFormatOfType(columns[prop].Type), colorIndex || ExcelKnownColors.ColorOfFFCC99, isBold));
                    }
                }
            }
        }
        else if (gridComponent.GetVisibleColumns && gridComponent.GetVisibleColumns instanceof Function) {
            //IWT ForaGrid
            let columns = gridComponent.GetVisibleColumns();
            this.Headers = [];
            for (var prop in columns) {
                if (columns.hasOwnProperty(prop)) {
                    this.AddHeader(columns[prop].Id, columns[prop].Text, new ExcelStyle(ExcelHelpers.GetExcelFormatOfType(columns[prop].ForaType), colorIndex || ExcelKnownColors.ColorOfFFCC99, isBold));
                }
            }
        }
        else
            WriteError("gridComponent minimum gereklilikleri karşılamıyor.");
    }

    AddDataGridRows(gridComponent, colorIndex, isBold = false) {
        if (!gridComponent)
            WriteError("gridComponent boş olmamalı ve Grid tipinde olmalı");
        let data;
        if (gridComponent.Rows && gridComponent.Rows instanceof Array) {
            //Interframe Grid
            data = gridComponent.Rows;
        }
        else if (gridComponent.GetRowsData && gridComponent.GetRowsData instanceof Function) {
            //IWT ForaGrid
            data = gridComponent.GetRowsData();
        }
        else
            WriteError("gridComponent minimum gereklilikleri karşılamıyor.");

        this.Cells = [];
        this.#c = 0;
        this.#r = 0;

        if (!data || data.length == 0) {
            return;
        }
        if (colorIndex === undefined) {
            colorIndex = null;
        }
        isBold = !!isBold;

        data.forEach(rowData => {
            this.Headers.forEach(header => {
                let text;
                // DisplayMemberName="BranchCode|||Number|||Suffix" şeklindeki kolonlar için özel kontrol eklendi.
                if (header.Id.indexOf("|||") > -1) {
                    text = header.Id.split("|||").map(r => r = rowData[r]).join("-");
                }
                else {
                    text = rowData[header.Id];
                }

                if (header.Format == ExcelCellFormatTypes.Date) {
                    text = text.replace(/\//g, '.');
                }
                this.AddCell(text, undefined, undefined, new ExcelStyle(header.Format, colorIndex, isBold))
            })
            this.AddNewRow();
        });
    }

    ImportDataGrid(gridComponent, headerColorIndex, isHeaderBold, rowsColorIndex, isRowsBold) {
        this.SetHeadersFromDataGrid(gridComponent, headerColorIndex, isHeaderBold);
        this.AddDataGridRows(gridComponent, rowsColorIndex, isRowsBold);
    }

    SetFromArray(data, showFromColumn, showToColumn) {
        this.Cells = [];
        if (fora.Inu(showFromColumn) || showFromColumn < 1)
            showFromColumn = 1;

        if (fora.Inu(showToColumn))
            showToColumn = Number.MAX_VALUE;

        if (showToColumn < showFromColumn)
            return;

        data.forEach(rowData => {
            for (let colIndex = showFromColumn - 1; colIndex < showToColumn && colIndex < rowData.length; colIndex++) {
                this.AddCell(rowData[colIndex]);
            }
            this.AddNewRow();
        });
    }
}
