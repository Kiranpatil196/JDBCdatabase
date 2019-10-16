
import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class ExcelDatabase {
   public static void main(String[] args) throws Exception {
      Class.forName("oracle.jdbc.driver.OracleDriver");
      Connection connect = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:XE", "system", "system");
      
      Statement stat= connect.createStatement();
      ResultSet resultSet = stat.executeQuery("select * from student3 order by rank asc ");
      HSSFWorkbook workbook = new HSSFWorkbook(); 
      HSSFSheet spreadsheet = workbook.createSheet("student");
      Row row = spreadsheet.createRow(1);
      Cell cell;
      cell = row.createCell(1);
      cell.setCellValue("Id");
      cell = row.createCell(2);
      cell.setCellValue("Name");
      cell = row.createCell(3);
      cell.setCellValue("Marks1");
      cell = row.createCell(4);
      cell.setCellValue("Marks2");
      cell = row.createCell(5);
      cell.setCellValue("Average");
      cell = row.createCell(6);
      cell.setCellValue("Rank");
      int i = 2;

      while(resultSet.next()) {
         row = spreadsheet.createRow(i);
         cell = row.createCell(1);
         cell.setCellValue(resultSet.getInt("id"));
         cell = row.createCell(2);
         cell.setCellValue(resultSet.getString("name"));
         cell = row.createCell(3);
         cell.setCellValue(resultSet.getInt("marks1"));
         cell = row.createCell(4);
         cell.setCellValue(resultSet.getInt("marks2"));
         cell = row.createCell(5);
         cell.setCellValue(resultSet.getInt("avg"));
         cell = row.createCell(6);
         cell.setCellValue(resultSet.getInt("rank"));
         i++;
      }
      row = spreadsheet.createRow(6);
      cell=row.createCell(2);
      cell.setCellValue("Total");
      
      cell=row.createCell(3);
      cell.setCellFormula("sum(d2:d6)");
      cell=row.createCell(4);
      cell.setCellFormula("sum(e2:e6)");
      //cell=row.createCell(5);
      //cell.setCellFormula("max(f2:f6)");
      
      row =spreadsheet.createRow(7);
      cell=row.createCell(2);
      cell.setCellValue("Maximum");
      cell=row.createCell(3);
    
      cell.setCellFormula("max(d2:d6)");
      cell=row.createCell(4);
      cell.setCellFormula("max(e2:e6)");
      
      row =spreadsheet.createRow(8);
      cell=row.createCell(2);
      cell.setCellValue("Minimum");
      cell=row.createCell(3);
    
      cell.setCellFormula("min(d2:d6)");
      cell=row.createCell(4);
      cell.setCellFormula("min(e2:e6)");
      
      String str="excel.xls";
      FileOutputStream out = new FileOutputStream(new File(str));
      workbook.write(out);
      out.close();
      System.out.println("excel.xls file written successfully !!!!!");
   }
}
