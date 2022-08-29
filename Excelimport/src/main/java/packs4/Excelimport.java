package packs4;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelimport {
	

	public static void main(String args[]) throws IOException
	{  
	  FileInputStream fis = new FileInputStream(new File("C:\\Users\\User\\OneDrive\\Documents\\Student Data.xlsx"));
	 
	  //workbook manipulation
	  XSSFWorkbook wb= new XSSFWorkbook(fis);  
	  
	  XSSFSheet sheet=wb.getSheetAt(0);  
	  //FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
	  for(Row row: sheet)     //iteration over row using for each loop  
	  {  
		  for(Cell cell : row){
			  System.out.print(cell+" ");
		  }
		  System.out.println();
		  System.out.println("test txt");
		  System.out.println("changes done");
	  }  
	}
	}
	  
	
	