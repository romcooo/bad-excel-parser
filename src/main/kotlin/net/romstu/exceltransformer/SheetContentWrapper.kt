package net.romstu.exceltransformer


/**
 * Sheet content wrapper that wraps the content of XSSFWorkbook.Sheet.
 *
 * @example of working with a sheet such as:
 * | name | surname | age |
 * | Mark | Vaughn  | 24  |
 * | John | Dean    | 45  |
 * | John | Gill    | 84  |
 * - filtering rows where column 'name'=="John":
 * sheet.contentMap.values.filter { it.fields["name"].equals("John") }
 * - adding a column "year of birth" with an approximate calculation
 * ```
 * sheet.headers.add("year of birth")
 * sheet.contentMap.entries.forEach {
 *     it.value.fields["year of birth"] = (Year.now().value - it.fields["age"]?.toInt).toString()
 * }
 * ```
 * @property sheetName name of the sheet
 * @property headers set of headers (column names), paired with their column index in the file
 * @property contentMap map of row index to [Row]
 * @constructor Create empty Sheet content wrapper
 */
class SheetContentWrapper(
    val sheetName: String
) {
    val headers: HashSet<Pair<String, Int>> = hashSetOf()
    val contentMap: HashMap<Int, Row> = hashMapOf()

    override fun toString(): String {
        return "SheetContentWrapper(sheetName='$sheetName', headers=$headers, columnMap=$contentMap)"
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