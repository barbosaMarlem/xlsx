var ByteArrayInputStream = Java.type('java.io.ByteArrayInputStream');
var XSSFWorkbook = Java.type('org.apache.poi.xssf.usermodel.XSSFWorkbook');
var CellReference = Java.type('org.apache.poi.ss.util.CellReference');
var DateUtil = Java.type('org.apache.poi.ss.usermodel.DateUtil');
var CellType = Java.type('org.apache.poi.ss.usermodel.CellType');

function read(file, metadata) {
    metadata = Object.assign({
        hasHeader: true
    }, metadata);

    var wb;

    if (file.length) {
        file = new ByteArrayInputStream(file);
    }

    wb = new XSSFWorkbook(file)
    var sheet = wb.getSheetAt(0);

    var resultRows = [];
    var headerNamesByIndex = {};

    sheet.rowIterator().forEachRemaining(function (row) {
        var isHeader = metadata.hasHeader && row.getRowNum() == 0;
        var rowObj = isHeader ? null : {};

        row.cellIterator().forEachRemaining(function (cell) {
            var columnIndex = cell.getColumnIndex();

            if (isHeader) {
                headerNamesByIndex[columnIndex] = getCellValue(cell);
                return;
            }

            var key = headerNamesByIndex[columnIndex];

            if (!key) {
                key = headerNamesByIndex[columnIndex] = CellReference.convertNumToColString(columnIndex);
            }

            rowObj[key] = getCellValue(cell, wb);
        });

        if (rowObj) {
            resultRows.push(rowObj);
        }
    });

    return resultRows;
}

function getCellValue(cell, wb) {
    if (cell == null) {
        return null;
    }

    switch (cell.getCellTypeEnum()) {
        case CellType.BLANK:
            return null;

        case CellType.BOOLEAN:
            return cell.getBooleanCellValue();

        case CellType.NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                var date = cell.getDateCellValue();

                if (date) {
                    return new Date(Number(date.getTime()));
                }

                return date;
            }
            
            return cell.getNumericCellValue();

        case CellType.STRING:
            return cell.getStringCellValue();

        case CellType.FORMULA:
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            return getCellValue(evaluator.evaluateInCell(cell));

        default:
            return null;
    }
}

exports = read;