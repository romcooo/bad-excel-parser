package net.romstu.excelparser

import mu.KotlinLogging
import java.util.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Excel reader
 * Used to convert an excel file into a list of sheets wrapped in SheetHolder-s.
 *
 * @property path
 * @constructor Create empty Excel reader
 */
class ExcelReader(
    private val path: String
) {
    private val logger = KotlinLogging.logger(this::class.simpleName!!)

    /**
     * Read file
     *
     * @param sheetNames - optional; if provided, reads only sheets matching (case-sensitive!) any of the sheet
     * names provided.
     * @return
     *  - If file could not be read, returns Optional.empty()
     *  - If file could be read correctly:
     *      - If no sheetNames provided, a list of all sheets from the file, each wrapped in a SheetHolder
     *      - If sheetNames provided, then only (wrapped) sheets from the file that match names with the ones provided.
     *      - If sheetNames provided but no sheetName matches, an Optional.of() with an empty list.
     */
    fun readFile(vararg sheetNames: String): Optional<List<SheetHolder>> {
        val wb = try {
            XSSFWorkbook(path)
        } catch (e: Exception) {
            logger.debug("Failed to open excel")
            e.printStackTrace()
            return Optional.empty()
        }
        val sheets = mutableListOf<SheetHolder>()
        wb.sheetIterator().forEach {
            if (sheetNames.isEmpty() || sheetNames.contains(it.sheetName)) {
                sheets.add(SheetHolder.fromSheet(it))
            }
        }
        wb.close()
        sheets.forEach { println(it) }
        return Optional.of(sheets)
    }
}
