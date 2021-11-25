package net.romstu.exceltransformer

import mu.KotlinLogging
import net.romstu.exceltransformer.SheetContentWrapper.Companion.fromSheet
import org.apache.poi.ss.usermodel.Sheet

/**
 * Sheet content wrapper that wraps the content of XSSFWorkbook.Sheet.
 * [fromSheet] can be used for creating an instance, given that the method's preconditions are met.
 *
 * @property sheetName - name of the sheet
 * @property headers - set of headers (column names), paired with their column index in the file
 * @property contentMap - map of row index to [Row]
 * @constructor Create empty Sheet content wrapper
 */
class SheetContentWrapper(
    val sheetName: String
) {
    val headers: HashSet<Pair<String, Int>> = hashSetOf()
    val contentMap: HashMap<Int, Row> = hashMapOf()

    companion object {
        private val logger = KotlinLogging.logger("SheetHolder companion")

        /**
         * "Factory method" static function that returns a [SheetContentWrapper] instance based on the provided [Sheet].
         * Assumes that the first/topmost row (index 0) represents the names of the columns, and maps this row onto the
         * [headers] hashset of pairs of <column index: Int, column name: String>.
         * The rest of the rows is then mapped onto the [contentMap] as a map of <row index, [Row]>.
         * Each row then further contains a map of <column name: String, value: String> (see [Row]).
         * @example an excel file containing something like
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
        fun fromSheet(sheet: Sheet): SheetContentWrapper {
            val holder = SheetContentWrapper(sheet.sheetName)
            for ((rowIndex, row) in sheet.rowIterator().withIndex()) {
                for ((columnIndex, cell) in row.iterator().withIndex()) {
                    cellValueAsString(cell).ifPresent { cellValue ->
                        if (rowIndex == 0) {
                            holder.headers.add(Pair(cellValue, columnIndex))
                        } else {
                            holder.headers.find { it.second == columnIndex }?.first?.let { headerName ->
                                holder.contentMap.getOrPut(rowIndex) { Row() }.fields[headerName] = cellValue
                            }
                        }
                        logger.debug("columnIndex=$columnIndex, rowIndex=$rowIndex, cell set to '$cellValue'")
                    }
                }
            }
            return holder
        }
    }

    override fun toString(): String {
        return "SheetHolder(sheetName='$sheetName', logger=$logger, headers=$headers, columnMap=$contentMap)"
    }
}

/**
 * Row
 * Represents the content of one excel row, with the key being the name of the column (see [SheetContentWrapper]).
 *
 * @property fields - a map of column/header name to actual value of the cell as a string.
 * @constructor Create empty Row
 */
data class Row(
    val fields: MutableMap<String, String> = mutableMapOf()
) {
    /**
     * Copy - "overrides" data class' default copy() method, which does not provide a deep copy with [fields].
     *
     * @return a deep copy of the instance on which it's called.
     */
    @Override
    fun copy(): Row {
        return Row(HashMap(this.fields))
    }
}