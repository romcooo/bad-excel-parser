package net.romstu.excelparser

import mu.KotlinLogging
import org.apache.poi.ss.usermodel.Sheet

class SheetHolderWriter(
    val fileSheet: Sheet,
    val sheetHolder: SheetHolder
) {
    private val logger = KotlinLogging.logger(this::class.simpleName!!)
    fun write() {
        writeHeaders()
        writeContent()
    }

    private fun writeHeaders() {
        val headerRow = fileSheet.createRow(0)
        for (header in sheetHolder.headerMap.entries) {
            headerRow.createCell(header.key)
                .setCellValue(header.value)
        }
    }

    private fun writeContent() {
        for (rowEntry in sheetHolder.contentMap.entries) {
            val currentRow = fileSheet.createRow(rowEntry.key)
            for (rowColumn in rowEntry.value.fields) {
                sheetHolder.headerMap.entries
                    .find { it.value == rowColumn.key }
                    ?.let {
                        currentRow.createCell(it.key).setCellValue(rowColumn.value)
                        logger.trace("rowKey = ${rowEntry.key}, columnKey=${it.key}, value = ${rowColumn.value}")
                    }
            }
        }
    }
}