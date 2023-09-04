import org.apache.commons.lang3.StringUtils
import org.apache.commons.lang3.math.NumberUtils
import org.apache.pdfbox.pdmodel.PDDocument
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.util.StringUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import technology.tabula.ObjectExtractor
import technology.tabula.PageIterator
import technology.tabula.Table
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm
import java.io.File
import java.io.FileOutputStream


fun main(args: Array<String>) {
    val wb = XSSFWorkbook()
    val dir = "C:\\Users\\Moldovan Zsombor\\Documents\\tarsasház"
    for (file in File(dir).listFiles()) {
        if (file.name.endsWith(".pdf")){
            processFile(wb, dir, StringUtils.substring(file.name, 0, -4))
        }
    }
    val fos = FileOutputStream("C:\\Users\\Moldovan Zsombor\\Documents\\tarsasház\\Naplófőkönyv.xlsx")
    wb.write(fos)
    fos.close()
}

private fun processFile(wb: XSSFWorkbook, baseDir: String, fileName: String) {
    val sheet = wb.createSheet(fileName)
    val pd: PDDocument = PDDocument.load(File("$baseDir\\$fileName.pdf"))
    val totalPages: Int = pd.getNumberOfPages()
    println("Total Pages in Document: $totalPages")
    val oe = ObjectExtractor(pd)
    val sea = SpreadsheetExtractionAlgorithm()
    val pages: PageIterator = oe.extract()
    // extract text from the table after detecting
    var rowNumber = 0
    for (page in pages) {
        val table: List<Table> = sea.extract(page)
        for (tables in table) {
            val rows = tables.getRows()
            for (i in rows.indices) {
                val cells = rows[i]
                val row = sheet.createRow(i + rowNumber)
                for (j in cells.indices) {
                    val cellValue = cells[j].getText()
                    if (NumberUtils.isCreatable(cellValue)) {
                        row.createCell(j, CellType.NUMERIC).setCellValue(NumberUtils.createDouble(cellValue))
                    } else {
                        row.createCell(j, CellType.STRING).setCellValue(cellValue)
                    }
                }
            }
            rowNumber += rows.size
        }
    }
}