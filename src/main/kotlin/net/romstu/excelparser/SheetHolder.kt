package net.romstu.excelparser

import mu.KotlinLogging
import org.apache.poi.ss.usermodel.Sheet

class SheetHolder(
    val sheetName: String
) {
    private val logger = KotlinLogging.logger(this::class.simpleName!!)
    val headerMap: HashMap<Int, String> = hashMapOf()
    val contentMap: HashMap<Int, Row> = hashMapOf()

    companion object {
        private val logger = KotlinLogging.logger(this::class.simpleName!!)
        fun fromSheet(sheet: Sheet): SheetHolder {
            val holder = SheetHolder(sheet.sheetName)
            for ((rowIndex, row) in sheet.rowIterator().withIndex()) {
                for ((columnIndex, cell) in row.iterator().withIndex()) {
                    cellValueAsString(cell).ifPresent { cellValue ->
                        if (rowIndex == 0) {
                            holder.headerMap[columnIndex] = cellValue
                        } else {
                            holder.headerMap[columnIndex]?.let { header ->
                                holder.contentMap.getOrPut(rowIndex) { Row() }.fields[header] = cellValue
                            }
                        }
                    }
                }
            }
            return holder
        }
    }

    override fun toString(): String {
        return "SheetHolder(sheetName='$sheetName', logger=$logger, headerMap=$headerMap, columnMap=$contentMap)"
    }
}

data class Row(
    val fields: MutableMap<String, String> = mutableMapOf()
) {
    @Override
    fun copy(): Row {
        return Row(HashMap(this.fields))
    }
}