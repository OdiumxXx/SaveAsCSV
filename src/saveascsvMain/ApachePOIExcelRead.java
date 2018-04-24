package saveascsvMain;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;

public class ApachePOIExcelRead {

    private static final String FILE_NAME = "Members List - 2015.xls";

    public static void main(String[] args) {

        try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new HSSFWorkbook(excelFile);
            int lastSheet = workbook.getNumberOfSheets()-1;
            Sheet datatypeSheet = workbook.getSheetAt(lastSheet);
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