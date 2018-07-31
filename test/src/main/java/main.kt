import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

val numCartesPorHoja = 35
val numCartesPorFila = 7

val numCartesPorHojaStandard = 18
val numCartesPorFilaStandard = 6


/**
 * Reads the value from the cell at the first row and first column of worksheet.
 */
fun readFromExcelFileEquipo() {
    val filepath = "./Kingdom Death Monster Cartas.xls"
    //Instantiate Excel workbook:
    val xlWbss = XSSFWorkbook()
    //Instantiate Excel worksheet:
    val xlWsss = xlWbss.createSheet()

    val inputStream = FileInputStream(filepath)
    //Instantiate Excel workbook using existing file:
    var xlWb = WorkbookFactory.create(inputStream)

    val filepathOut = "outEquipo.xls"

    //Row index specifies the row in the worksheet (starting at 0):
    var rowNumber = 0
    //Cell index specifies the column within the chosen row (starting at 0):
    val columnNumber = 0

    //Get reference to first sheet:
    val xlWs = xlWb.getSheetAt(0)

    while (rowNumber < 840) {

        val text = xlWs.getRow(rowNumber).getCell(columnNumber).stringCellValue
        print("${rowNumber+1} - $text  con % = ${(rowNumber-1) % (numCartesPorFila*2)} -- pasa a ser: ")
        if ((rowNumber-1) % (numCartesPorHoja*2) >= numCartesPorHoja) {
            if ((rowNumber - 1) % numCartesPorFila == 0) {
                xlWsss.createRow(rowNumber + 6).createCell(0).setCellValue(text)
                print(" ${rowNumber + 6} \n")

            } else if ((rowNumber - 1) % numCartesPorFila == 1) {
                xlWsss.createRow(rowNumber + 4).createCell(0).setCellValue(text)
                print(" ${rowNumber + 4} \n")

            } else if ((rowNumber - 1) % numCartesPorFila == 2) {
                xlWsss.createRow(rowNumber + 2).createCell(0).setCellValue(text)
                print(" ${rowNumber + 2} \n")

            } else if ((rowNumber - 1) % numCartesPorFila == 3) {
                xlWsss.createRow(rowNumber).createCell(0).setCellValue(text)
                print(" ${rowNumber} \n")

            } else if ((rowNumber - 1) % numCartesPorFila == 4) {
                xlWsss.createRow(rowNumber - 2).createCell(0).setCellValue(text)
                print(" ${rowNumber - 2} \n")

            } else if ((rowNumber - 1) % numCartesPorFila == 5) {
                xlWsss.createRow(rowNumber - 4).createCell(0).setCellValue(text)
                print(" ${rowNumber - 4} \n")

            } else if ((rowNumber - 1) % numCartesPorFila == 6) {
                xlWsss.createRow(rowNumber - 6).createCell(0).setCellValue(text)
                print(" ${rowNumber - 6} \n")

            } else
                print("\n")
        } else {
            xlWsss.createRow(rowNumber).createCell(0).setCellValue(text)
            print(" EL MISMO\n")
        }

        rowNumber++
    }
    val outputStream = FileOutputStream(filepathOut)
    xlWbss.write(outputStream)
    xlWbss.close()
}

/**
 * Reads the value from the cell at the first row and first column of worksheet.
 */
fun readFromExcelFile( ) {
    val filepath = "./Kingdom Death Monster Cartas Standard.xls"
    //Instantiate Excel workbook:
    val xlWbss = XSSFWorkbook()
    //Instantiate Excel worksheet:
    val xlWsss = xlWbss.createSheet()

    val inputStream = FileInputStream(filepath)
    //Instantiate Excel workbook using existing file:
    var xlWb = WorkbookFactory.create(inputStream)

    val filepathOut = "out.xls"

    //Row index specifies the row in the worksheet (starting at 0):
    var rowNumber = 0
    //Cell index specifies the column within the chosen row (starting at 0):
    val columnNumber = 0

    //Get reference to first sheet:
    val xlWs = xlWb.getSheetAt(0)

    while (rowNumber < 1477) {

        val text = xlWs.getRow(rowNumber).getCell(columnNumber).stringCellValue
        print( "  ${rowNumber+1} - $text con % = ${(rowNumber-1) % numCartesPorFilaStandard} -- pasa a ser: ")
        if ((rowNumber-1) % (numCartesPorHojaStandard*2) >= numCartesPorHojaStandard) {
            if ((rowNumber - 1) % numCartesPorFilaStandard == 0) {
                xlWsss.createRow(rowNumber + 5).createCell(0).setCellValue(text)
                print(" ${rowNumber + 5}")
            } else if ((rowNumber - 1) % numCartesPorFilaStandard == 1) {
                xlWsss.createRow(rowNumber + 3).createCell(0).setCellValue(text)
                print(" ${rowNumber + 3}")
            } else if ((rowNumber - 1) % numCartesPorFilaStandard == 2) {
                xlWsss.createRow(rowNumber + 1).createCell(0).setCellValue(text)
                print(" ${rowNumber + 1}")
            } else if ((rowNumber - 1) % numCartesPorFilaStandard == 3) {
                xlWsss.createRow(rowNumber - 1).createCell(0).setCellValue(text)
                print(" ${rowNumber - 1}")
            } else if ((rowNumber - 1) % numCartesPorFilaStandard == 4) {
                xlWsss.createRow(rowNumber - 3).createCell(0).setCellValue(text)
                print(" ${rowNumber - 3}")
            } else if ((rowNumber - 1) % numCartesPorFilaStandard == 5) {
                xlWsss.createRow(rowNumber - 5).createCell(0).setCellValue(text)
                print(" ${rowNumber - 5}")
            }
        } else {
            xlWsss.createRow(rowNumber).createCell(0).setCellValue(text)
            print("LA MISMA")
        }
        print("\n")
        rowNumber++
    }
    val outputStream = FileOutputStream(filepathOut)
    xlWbss.write(outputStream)
    xlWbss.close()
}

fun main(args: Array<String>) {
    readFromExcelFile()
    //readFromExcelFileEquipo()
}

