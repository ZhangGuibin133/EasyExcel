import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

object Excel {
    //从目标文件中获取表名对应List<List<String>> Map,一行为一个List<String>
    fun loadSheets(file: File): Map<String, List<List<String>>> {
        //从文件流获取内容
        val map = mutableMapOf<String, List<List<String>>>()
        fun Row.getCells() = (0 until lastCellNum).map { getCell(it)?.toString() ?: "" }
        FileInputStream(file).use { fis ->
            WorkbookFactory.create(fis).use { book ->
                book.forEach { sheet -> map[sheet.sheetName] = sheet.map { row -> row.getCells() } }
            }
        }
        return map.toMap()
    }

    //获取目标文件名 获取sheet1的内容
    fun loadSheet(file: File) = loadSheets(file).values.first() ?: emptyList()

    //将表名相应List对应保存为sheet
    fun saveSheets(file: File, map: Map<String, List<List<String>>>) {
        //将字符串依次写入行 如果全数字,设为数字格式
        XSSFWorkbook().use { book ->
            //按单元格数量创建cell,如内容为数字文本,设为数字值,否则为文本
            fun Row.setCells(cells: List<String>) {
                cells.indices.forEach { i ->
                    cells[i].toDoubleOrNull()?.run {
                        createCell(i, CellType.NUMERIC).setCellValue(this)
                    } ?: createCell(i, CellType.STRING).setCellValue(cells[i])
                }
            }
            map.forEach {
                val sheet = book.createSheet(it.key)
                it.value.forEachIndexed { rowNumber, cells ->
                    sheet.createRow(rowNumber).setCells(cells)
                }
            }
            FileOutputStream(file).use { book.write(it) }
        }
    }

    fun saveSheet(file: File, sheet: List<List<String>>, sheetName: String = "sheet1") {
        saveSheets(file, mapOf(sheetName to sheet))
    }

    //获取目标文件名 一行一个List,获取指定表头内容 文件名对应一个List<Line>
    fun loadSheetsWithFieldName(file: File, fieldNames: List<String>): List<List<String>> {
        val filteredLines = mutableListOf<List<String>>()
        FileInputStream(file).use { fis ->
            WorkbookFactory.create(fis).use {
                for (sheet in it) {
                    val columnHead = sheet.first().run { (0 until lastCellNum).map { getCell(it)?.toString() ?: "" } }
                    val columnIndex = fieldNames.map { columnHead.indexOf(it) }.filter { it >= 0 }
                    sheet.drop(1).map { row ->
                        columnIndex.map { "${row.getCell(it) ?: ""}" }
                    }.let { filteredLines.addAll(it) }
                }
            }
        }
        return filteredLines.toList()
    }
}
