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
  //private static final String FILE_NAME = "members.xls"; 
  private static int ourSheet = 75;

  // Our Text Area Editor (So we don't have to keep referring to saveascsvMain.etc.etc.etc) 
  private static void textAreaEditor(String s) {
    saveascsvMain.OpenDialog.textArea.append(s);
  }

  // Start the Conversion!
  public static void main(String filePath) {

    try {

      if (filePath == null) {
        textAreaEditor("\n  ERROR: You have not selected a spreadsheet.\n");
      }
      FileInputStream excelFile = new FileInputStream(new File(filePath));
      Workbook workbook = new XSSFWorkbook(excelFile);      
      int sheetsNumberOf = workbook.getNumberOfSheets();  

      //Setup regex to find the spreadsheet we want to work with
      Pattern p = Pattern.compile("Consolidate");   // the pattern to search for
      System.out.println("- Checking all sheets in workbook for one with 'Consolidate' in the title....");
      textAreaEditor("- Checking all sheets in workbook for one with 'Consolidate' in the title....\n");
      // for each sheet, search for the regex pattern 'p' if the matcher makes a find of that pattern, set that to be 'ourSheet' and tell everyone about it!
      for(int i = 0; i < sheetsNumberOf; i++) {   
        String sheetName = workbook.getSheetName(i);
        Matcher m = p.matcher(sheetName);  
        if (m.find()) {
          System.out.println("    Found Spreadsheet with phrase 'Consolidate'");
          textAreaEditor("    Found Spreadsheet with phrase 'Consolidate'\n");
          ourSheet = i;
        }  else {
          System.out.println("    Checking sheet "+i+"...");
          textAreaEditor("    Checking sheet "+i+"...\n");
        }
      }


      // Now we have the spreadsheet to work with, it's time to iterate through it and create our CSV
      Sheet datatypeSheet = workbook.getSheetAt(ourSheet);
      Iterator<Row> iterator = datatypeSheet.iterator();
      PrintWriter pw = new PrintWriter(new File("memdraw.csv"));
      StringBuilder sb = new StringBuilder();
      int Entries = 0;
      textAreaEditor("- Consolidate Members Sheet found!\n- Creating CSV...\n---------------------------------------\n");

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
            textAreaEditor(cellValue + ",");            
            if (cellValue.contains("\n") == true) {
              cellValue = "Cell Value Contains NEWLINE - Fix your Spreadsheet!";
              System.out.print("*** REPLACED A NEW LINE ***");
              textAreaEditor("\n*** WARNING: Cell Value Contains a linebreak. I have removed it for the CSV file, but you should fix your Spreadsheet! ***\n");
            }
            sb.append(cellValue + ",");            
          } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {                        
            int cellValueNum = (int) currentCell.getNumericCellValue();
            System.out.print(cellValueNum + ",");
            textAreaEditor(cellValueNum + ",");
            sb.append(cellValueNum + ",");            
          }          
        }
        // END ROW
        Entries++;         
        textAreaEditor("\n");
        sb.append("\n");
      }
      // END CSV
      pw.write(sb.toString());
      pw.close();
      workbook.close();

      // ADDITIONAL APPLET OUTPUT
      Entries--;
      textAreaEditor("------------------------------\n");
      textAreaEditor("- Succesfully created memdraw.csv!\n");      
      textAreaEditor("- Members added: "+Entries+"\n");
      textAreaEditor("\n");
      textAreaEditor("Done!\n");
      textAreaEditor("You may now close this window\n");
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }

  }
}