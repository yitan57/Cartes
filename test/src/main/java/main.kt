import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

/**
 * Writes the value "TEST" to the cell at the first row and first column of worksheet.
 */
fun writeToExcelFile(text: String, filepath: String, rowNumber: Int) {
    //Instantiate Excel workbook:
    val xlWb = XSSFWorkbook()
    //Instantiate Excel worksheet:
    val xlWs = xlWb.createSheet()

    //Write text value to cell located at ROW_NUMBER / COLUMN_NUMBER:
    xlWs.createRow(5).createCell(0).setCellValue(text)

    //print(text + "  $rowNumber \n")

    //Write file:
    val outputStream = FileOutputStream(filepath)
    xlWb.write(outputStream)
    xlWb.close()
}


/**
 * Reads the value from the cell at the first row and first column of worksheet.
 */
fun readFromExcelFileEquipo(filepath: String) {
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
        print( text + "  $rowNumber con % = ${(rowNumber-1) % 7} -- pasa a ser: ")
        if (text.contains("Trasera")) {
            if ((rowNumber - 1) % 7 == 0) {
                xlWsss.createRow(rowNumber + 6).createCell(0).setCellValue(text)
                print(" ${rowNumber + 6} \n")

            } else if ((rowNumber - 1) % 7 == 1) {
                xlWsss.createRow(rowNumber + 4).createCell(0).setCellValue(text)
                print(" ${rowNumber + 4} \n")

            } else if ((rowNumber - 1) % 7 == 2) {
                xlWsss.createRow(rowNumber + 2).createCell(0).setCellValue(text)
                print(" ${rowNumber + 2} \n")

            } else if ((rowNumber - 1) % 7 == 3) {
                xlWsss.createRow(rowNumber).createCell(0).setCellValue(text)
                print(" ${rowNumber} \n")

            } else if ((rowNumber - 1) % 7 == 4) {
                xlWsss.createRow(rowNumber - 2).createCell(0).setCellValue(text)
                print(" ${rowNumber - 2} \n")

            } else if ((rowNumber - 1) % 7 == 5) {
                xlWsss.createRow(rowNumber - 4).createCell(0).setCellValue(text)
                print(" ${rowNumber - 4} \n")

            } else if ((rowNumber - 1) % 7 == 6) {
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
fun readFromExcelFile(filepath: String) {
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
        print( text + "  $rowNumber con % = ${(rowNumber-1) % 6} -- pasa a ser: ")
        if (text.contains("TRASERAS")) {
            if ((rowNumber - 1) % 6 == 0) {
                xlWsss.createRow(rowNumber + 5).createCell(0).setCellValue(text)
                print(" ${rowNumber + 5} \n")
            } else if ((rowNumber - 1) % 6 == 1) {
                xlWsss.createRow(rowNumber + 3).createCell(0).setCellValue(text)
                print(" ${rowNumber + 3} \n")
            } else if ((rowNumber - 1) % 6 == 2) {
                xlWsss.createRow(rowNumber + 1).createCell(0).setCellValue(text)
                print(" ${rowNumber + 1} \n")
            } else if ((rowNumber - 1) % 6 == 3) {
                xlWsss.createRow(rowNumber - 1).createCell(0).setCellValue(text)
                print(" ${rowNumber - 1} \n")
            } else if ((rowNumber - 1) % 6 == 4) {
                xlWsss.createRow(rowNumber - 3).createCell(0).setCellValue(text)
                print(" ${rowNumber - 3} \n")
            } else if ((rowNumber - 1) % 6 == 5) {
                xlWsss.createRow(rowNumber - 5).createCell(0).setCellValue(text)
                print(" ${rowNumber - 5} \n")
            }
        } else {
            xlWsss.createRow(rowNumber).createCell(0).setCellValue(text)
        }

        rowNumber++
    }
    val outputStream = FileOutputStream(filepathOut)
    xlWbss.write(outputStream)
    xlWbss.close()
}

fun main(args: Array<String>) {
    //val filepath = "./Kingdom Death Monster Cartas Standard.xls"
    val filepath = "./Kingdom Death Monster Cartas.xls"
    //writeToExcelFile(filepath)
    //readFromExcelFile(filepath)
    readFromExcelFileEquipo(filepath)
}

