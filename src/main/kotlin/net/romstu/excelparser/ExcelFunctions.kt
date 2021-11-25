package net.romstu.excelparser

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import java.util.*


/**
 * Write to new file
 * Takes a list of sheets wrapped in SheetHolder-s, and creates a new Excel file containing those sheets.
 *
 * @param sheets
 * @param newFilePath
 */
fun writeToNewFile(sheets: List<SheetHolder>, newFilePath: String) {
    val wb = XSSFWorkbook()
    for (sheetHolder in sheets) {
        SheetHolderWriter(
            fileSheet = wb.createSheet(sheetHolder.sheetName),
            sheetHolder = sheetHolder
        ).write()
    }
    wb.write(FileOutputStream(newFilePath))
    wb.close()
}

/**
 * Cell value as string
 * Returns an Optional<String> containing the cell value converted to a string, no matter the original cell type.
 *
 * @param cell to be "read"
 * @return Optional.of(cell value converted to string), or Optional.empty if cellType==ERROR
 */
fun cellValueAsString(cell: Cell): Optional<String> {
    return when (cell.cellType) {
        CellType.STRING -> Optional.of(cell.stringCellValue)
        CellType.NUMERIC -> Optional.of(cell.numericCellValue.toString())
        CellType.BLANK -> Optional.of("")
        CellType._NONE -> Optional.of("")
        CellType.FORMULA -> Optional.of(cell.cellFormula.toString())
        CellType.BOOLEAN -> Optional.of(cell.booleanCellValue.toString())
        CellType.ERROR -> Optional.empty()
        else -> Optional.empty()
    }
}