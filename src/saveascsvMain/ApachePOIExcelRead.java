package saveascsvMain;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ApachePOIExcelRead {
  //  private static final String FILE_NAME = "members.xls"; 

  private static int ourSheet = 75;

  public static void main(String filePath) {

    try {

      FileInputStream excelFile = new FileInputStream(new File(filePath));
      Workbook workbook = new XSSFWorkbook(excelFile);      
      int sheetsNumberOf = workbook.getNumberOfSheets();  

      //Setup regex to find the sheet we want to work with
      Pattern p = Pattern.compile("Consolidate");   // the pattern to search for

      // for each sheet, search for the regex pattern 'p' if the matcher makes a find of that pattern, set that to be oursheet and tell everyone about it!
      for(int i = 0; i < sheetsNumberOf; i++) {   
        String sheetName = workbook.getSheetName(i);
        Matcher m = p.matcher(sheetName);  
        if (m.find()) {
          System.out.println("FOUND SHEET CONSOLIDATE!");
          ourSheet = i;
        }  else {
          System.out.println("Not this sheet");
        }
      }






      Sheet datatypeSheet = workbook.getSheetAt(ourSheet);
      Iterator<Row> iterator = datatypeSheet.iterator();

      PrintWriter pw = new PrintWriter(new File("memdraw.csv"));
      StringBuilder sb = new StringBuilder();

      while (iterator.hasNext()) {

        Row currentRow = iterator.next();
        Iterator<Cell> cellIterator = currentRow.iterator();

        while (cellIterator.hasNext()) {

          Cell currentCell = cellIterator.next();
          //getCellTypeEnum deprecated for version 3.15
          //getCellTypeEnum will be renamed to getCellType in version 4.0
          if (currentCell.getCellTypeEnum() == CellType.STRING) {
            String cellValue = currentCell.getStringCellValue();
            System.out.print(cellValue + ","); 
            if (cellValue.contains("\n") == true) {
              cellValue = "Cell Value Contains NEWLINE - Fix your Spreadsheet!";
              System.out.print("*** REPLACED A NEW LINE ***");
            }
            sb.append(cellValue + ",");
          } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {                        
            int cellValueNum = (int) currentCell.getNumericCellValue();
            System.out.print(cellValueNum + ",");
            sb.append(cellValueNum + ",");
          }

        }
        //                System.out.println("NEWLINE " + datatypeSheet.getLastRowNum());
        System.out.println("");  
        sb.append("\n");


      }
      pw.write(sb.toString());
      pw.close();
      workbook.close();
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }

  }
}