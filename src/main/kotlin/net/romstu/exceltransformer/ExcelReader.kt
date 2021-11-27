package net.romstu.exceltransformer

import mu.KotlinLogging
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.util.*

/**
 * Excel reader
 * Used to convert an Excel file into a list of sheets wrapped in [SheetContentWrapper]-s.
 *
 * @property path path to the file to be read
 * @property wrapSheet function that accepts a [Sheet] and returns the contents wrapped in a [SheetContentWrapper].
 * Default implementation is [wrapSheetDefault]
 * @constructor Create empty Excel reader
 */
class ExcelReader(
    private val path: String,
    private val wrapSheet: (sheet: Sheet) -> SheetContentWrapper = ::wrapSheetDefault
) {
    private val logger = KotlinLogging.logger(this::class.simpleName ?: "ExcelReader")

    /**
     * Read file
     *
     * @param sheetNames - optional; if provided, reads only sheets matching (case-sensitive!) any of the sheet
     * names provided.
     * @return
     *  - If file could not be read, returns Optional.empty()
     *  - If file could be read correctly:
     *      - If no sheetNames provided, a list of all sheets from the file, each wrapped in a [SheetContentWrapper]
     *      - If sheetNames provided, then only (wrapped) sheets from the file that match names with the ones provided.
     *      - If sheetNames provided but no sheetName matches, an Optional.of() with an empty list.
     */
    fun readFile(vararg sheetNames: String): Optional<List<SheetContentWrapper>> {
        val wb = try {
            XSSFWorkbook(path)
        } catch (e: Exception) {
            logger.warn("Failed to open excel file at $path")
            e.printStackTrace()
            return Optional.empty()
        }
        val sheets = mutableListOf<SheetContentWrapper>()
        wb.sheetIterator().forEach {
            if (sheetNames.isEmpty() || sheetNames.contains(it.sheetName)) {
                sheets.add(wrapSheet(it))
            }
        }
        wb.close()
        logger.debug("Read files: $sheets")
        return Optional.of(sheets)
    }
}

/**
 * Default mapping/wrapping function that returns a [SheetContentWrapper] instance based on the provided [Sheet].
 * Assumes that the first/topmost row (index 0) represents the names of the columns, and maps this row onto the
 * [SheetContentWrapper.headers] hashset of pairs of <column index: Int, column name: String>.
 * The rest of the rows is then mapped onto the [SheetContentWrapper.contentMap] as a map of <row index, [Row]>.
 * Each row then further contains a map of <column name: String, value: String> (see [Row]).
 * @example an Excel file containing something like
 * | name | surname | age |
 * | Mark | Vaughn  | 24  |
 * | John | Dean    | 45  |
 * would result in headers=[(1, name), (2, surname), (3, age)]
 * and contentMap=[
 *      1=Row(fields={"name" to "Mark", "surname" to "Vaughn", "age" to "24"}),
 *      2=Row(fields={"name" to "John", "surname" to "Dean", "age" to "45"}]
 *
 * @param sheet
 * @return
 */
fun wrapSheetDefault(sheet: Sheet): SheetContentWrapper {
    val logger = KotlinLogging.logger("defaultWrapSheet")
    val wrapper = SheetContentWrapper(sheet.sheetName)
    for ((rowIndex, row) in sheet.rowIterator().withIndex()) {
        for ((columnIndex, cell) in row.iterator().withIndex()) {
            cellValueAsString(cell).ifPresent { cellValue ->
                if (rowIndex == 0) {
                    wrapper.headers.add(Pair(cellValue, columnIndex))
                } else {
                    wrapper.headers.find { it.second == columnIndex }?.first?.let { headerName ->
                        wrapper.contentMap.getOrPut(rowIndex) { Row() }.fields[headerName] = cellValue
                    }
                }
                logger.debug("columnIndex=$columnIndex, rowIndex=$rowIndex, cell set to '$cellValue'")
            }
        }
    }
    return wrapper
}
