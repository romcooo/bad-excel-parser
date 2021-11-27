package net.romstu.exceltransformer

import mu.KotlinLogging
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import java.io.FileOutputStream

/**
 * Sheet wrapper writer
 *
 * @property workbook - provide a workbook to which sheets will be added and which will be used to create an excel file.
 * @constructor Create a Sheet wrapper writer.
 */
class SheetContentWrapperWriter(
    val workbook: Workbook
) {
    private val logger = KotlinLogging.logger(this::class.simpleName ?: "SheetContentWrapperWriter")
    private var closed = false

    /**
     * Add sheet to the workbook.
     *
     * @param sheetContentWrapper - to be added to the workbook as a sheet.
     */
    fun addSheet(sheetContentWrapper: SheetContentWrapper): Boolean {
        if (closed) {
            logger.warn("workbook is already closed!")
            return false
        }
        val fileSheet = try {
            workbook.createSheet(sheetContentWrapper.sheetName)
        } catch (e: IllegalArgumentException) {
            logger.warn(e.message)
            return false
        }
        writeHeaders(sheetContentWrapper, fileSheet)
        writeContent(sheetContentWrapper, fileSheet)
        return true
    }

    /**
     * Write sheets to file and close the workbook.
     * Cannot be reused after closing.
     *
     * @param filePath - path where the file should be written.
     */
    fun writeToFileAndClose(filePath: String) {
        workbook.write(FileOutputStream(filePath))
        closed = true
        workbook.close()
    }

    private fun writeHeaders(sheetContentWrapper: SheetContentWrapper, fileSheet: Sheet) {
        val headerRow = fileSheet.createRow(0)
        for (header in sheetContentWrapper.headers) {
            headerRow.createCell(header.second)
                .setCellValue(header.first)
        }
    }

    private fun writeContent(sheetContentWrapper: SheetContentWrapper, fileSheet: Sheet) {
        for (rowEntry in sheetContentWrapper.contentMap.entries) {
            val currentRow = fileSheet.createRow(rowEntry.key)
            for (rowColumn in rowEntry.value.fields) {
                sheetContentWrapper.headers
                    .find { it.first == rowColumn.key }
                    ?.let {
                        currentRow.createCell(it.second).setCellValue(rowColumn.value)
                        logger.trace("Writing: rowKey = ${rowEntry.key}, columnKey=${it.second}, value = ${rowColumn.value}")
                    }
            }
        }
    }
}