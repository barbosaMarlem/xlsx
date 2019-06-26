var ByteArrayOutputStream = Java.type('java.io.ByteArrayOutputStream');
var CellType = Java.type('org.apache.poi.ss.usermodel.CellType');
var HorizontalAlignment = Java.type('org.apache.poi.ss.usermodel.HorizontalAlignment');
var VerticalAlignment = Java.type('org.apache.poi.ss.usermodel.VerticalAlignment');
var XSSFFont = Java.type('org.apache.poi.xssf.usermodel.XSSFFont');
var XSSFWorkbook = Java.type('org.apache.poi.xssf.usermodel.XSSFWorkbook');
var IndexedColors = Java.type('org.apache.poi.ss.usermodel.IndexedColors');
var FillPatternType = Java.type('org.apache.poi.ss.usermodel.FillPatternType');

var DEFAULT_DATE_FORMAT = 'DD/MM/yyyy';
var DEFAULT_TIME_FORMAT = 'HH:MM';
var DEFAULT_DATETIME_FORMAT = DEFAULT_DATE_FORMAT + ' ' + DEFAULT_TIME_FORMAT;

var DEFAULT_CURRENCY_FORMAT = 'R$ #,##0.00';

var FONT_OPTION_NAMES = ['fontName', 'fontSize', 'fontColor', 'bold', 'italic', 'striked', 'underline', 'doubleUnderline'];
var FONTS_CACHE = {};
var STYLES_CACHE = {};

function create(rows, metadata) {
  
  FONTS_CACHE = {};
  STYLES_CACHE = {};
  
  metadata = Object.assign({
    asByteArray: true,
    autoSize: true,
    hasHeader: true
  }, metadata);

  var wb;

  try {
    wb = new XSSFWorkbook();

    var createHelper = wb.getCreationHelper();
    var sheet = wb.createSheet();

    var sheetProperties = metadata.sheetProperties;
    var columnsMD = metadata.columns;
    var generalStyle = metadata.style;
    var headerStyleMD = metadata.headerStyle;
    var rowStyle = metadata.rowStyle;

    if (sheetProperties) {
      if (sheetProperties.password) {
        sheet.protectSheet(sheetProperties.password);
        sheet.lockDeleteColumns(true);
        sheet.lockInsertColumns(true);
      }

      if (sheetProperties.name) {
        wb.setSheetName(0, sheetProperties.name);
      }
    }

    rows.forEach(function(row, index) {
      var sheetHeader = metadata.hasHeader && index == 0 ? sheet.createRow(0) : null;
      var sheetRow = sheet.createRow(metadata.hasHeader ? index + 1 : index);
      var rowAux = metadata.orderByColumnConfig ? columnsMD : row;

      Object.keys(rowAux).forEach(function(key, cellIndex) {
        var value = row[key];
        var valueClass = value && value.constructor.name;

        var columnMD = columnsMD && columnsMD[key];

        if (rowStyle && rowStyle.height) {
          sheetRow.setHeight(rowStyle.height);
        }

        if (sheetHeader != null) {

          if (headerStyleMD && headerStyleMD.height) {
            sheetHeader.setHeight(headerStyleMD.height);
          }

          var style = Object.assign({}, generalStyle, columnMD, headerStyleMD);

          var headerCell = createCell(wb, sheetHeader, cellIndex, style, createHelper, valueClass);

          if (columnMD && columnMD.description) {
            headerCell.setCellValue(columnMD.description);
          } else {
            headerCell.setCellValue(key);
          }
        }

        var style = Object.assign({}, generalStyle, columnMD);
        var cell = createCell(wb, sheetRow, cellIndex, style, createHelper, valueClass);

        var formula = (columnMD && columnMD.formula) ? columnMD.formula : null;
        setTypedValue(cell, value, valueClass, formula);
      });
    });

    if (metadata.autoSize) {
      var headerRow = sheet.getRow(0);
      headerRow.cellIterator().forEachRemaining(function(cell) {
        sheet.autoSizeColumn(cell.getColumnIndex());
      });
    } else {
      var headerRow = sheet.getRow(0);
      headerRow.cellIterator().forEachRemaining(function(cell) {
        var columnsKeys = Object.keys(columnsMD);
        var columnIndex = cell.getColumnIndex()

        if (columnsMD && columnsKeys.length > 0) {
          var columnMD = columnsMD[columnsKeys[columnIndex]]
          if (columnMD.size && columnMD.size.toString().toUpperCase() === 'AUTO') {
            sheet.autoSizeColumn(cell.getColumnIndex());
          } else if (!isNaN(columnMD.size) && columnMD.size >= 0) {
            sheet.setColumnWidth(cell.getColumnIndex(), columnMD.size);
          }
        }
      });
    }

    if (metadata.asByteArray) {
      var baos = new ByteArrayOutputStream();
      wb.write(baos);
      return baos.toByteArray();
    } else {
      return wb;
    }
  } finally {
    if (wb) {
      wb.close();
    }
  }
}

function createCell(wb, sheetRow, cellIndex, style, createHelper, valueClass) {
  var cell = sheetRow.createCell(cellIndex);

  style.format = resolveDataFormat(style, valueClass);

  if (Object.keys(style).length) {
    var cacheKey = JSON.stringify(style);
    var cellStyle = STYLES_CACHE[cacheKey];

    if (!cellStyle) {
      cellStyle = wb.createCellStyle();

      updateCellStyle(wb, cellStyle, style, createHelper);

      STYLES_CACHE[cacheKey] = cellStyle;
    }

    cell.setCellStyle(cellStyle);
  }

  return cell;
}

function updateCellStyle(wb, cellStyle, options, createHelper) {
  if (!options) {
    return;
  }

  if (options.backgroundColor) {
    cellStyle.setFillForegroundColor(IndexedColors[options.backgroundColor.toUpperCase()].getIndex());
    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
  }

  if (options.format) {
    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(options.format));
  }

  if(options.unlocked) {
    cellStyle.setLocked(false);
  }

  var font = resolveFont(wb, options);

  if (font) {
    cellStyle.setFont(font);
  }

  if (options.horizontalAlignment) {
    cellStyle.setAlignment(HorizontalAlignment[options.horizontalAlignment.toUpperCase()]);
  }
  if (options.verticalAlignment) {
    cellStyle.setVerticalAlignment(VerticalAlignment[options.verticalAlignment.toUpperCase()]);
  }
}

function resolveFont(wb, options) {
  if (!options) {
    return undefined;
  }

  var fontOpts = FONT_OPTION_NAMES.reduce(function(map, optName) {
    if (options[optName]) {
      map[optName] = options[optName];
    }

    return map;
  }, {});

  var font;

  if (Object.keys(fontOpts).length) {
    var key = JSON.stringify(fontOpts);

    font = FONTS_CACHE[key];

    if (!font) {
      font = wb.createFont();

      font.setFontName(fontOpts.fontName || XSSFFont.DEFAULT_FONT_NAME);
      font.setFontHeightInPoints(fontOpts.fontSize || XSSFFont.DEFAULT_FONT_SIZE);

      if (fontOpts.bold) {
        font.setBold(fontOpts.bold);
      }

      if (fontOpts.italic) {
        font.setItalic(fontOpts.italic);
      }

      if (fontOpts.striked) {
        font.setStrikeout(fontOpts.striked);
      }

      if (fontOpts.underline) {
        font.setUnderline(0x21);
      }

      if (fontOpts.doubleUnderline) {
        font.setUnderline(0x22);
      }

      if (fontOpts.fontColor) {
        font.setColor(IndexedColors[fontOpts.fontColor.toUpperCase()].getIndex());
      }

      FONTS_CACHE[key] = font;
    }
  }

  return font;
}

function resolveDataFormat(options, valueClass) {
  if (options.format) {
    return options.format;
  }

  switch (options.type) {
    case 'currency':
      return DEFAULT_CURRENCY_FORMAT;
    case 'time':
      return DEFAULT_TIME_FORMAT;
    case 'date':
      return DEFAULT_DATE_FORMAT;
    case 'datetime':
      return DEFAULT_DATETIME_FORMAT;
    default:
      var isDate = valueClass === 'Date';
      return isDate ? DEFAULT_DATETIME_FORMAT : null;
  }
}

function setTypedValue(cell, value, valueClass, formula) {

  if (value === null || value === undefined) {
    cell.setCellType(CellType.BLANK);
    return;
  }

  if (valueClass == 'Date') {
    value = new java.util.Date(value.getTime());
  }

  if (formula) {
    cell.setCellType(CellType.FORMULA);
    cell.setCellFormula(value);
  } else {
    cell.setCellValue(value);
  }

}

exports = create;
