package excel;
import java.io.*;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;

public class Excelwrite 
{
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		
		//Blank workbook
		        XSSFWorkbook workbook = new XSSFWorkbook(); 
		         
		        //Create a blank sheet
		        XSSFSheet sheet = workbook.createSheet("Students");
		            
		        ArrayList<String> studentList = new ArrayList<String>();
		        studentList.add("Student Name");
		        studentList.add("Rincy");
		        studentList.add("Layana");
		        studentList.add("Anil");
		        studentList.add("Athulya");
		for (int counter = 0; counter < studentList.size(); counter++) 
		        { 		      
		            
		        	 Row row = sheet.createRow(counter);
		        	 Cell cell = row.createCell(0);
		        	 cell.setCellValue(studentList.get(counter));        			
		        }
		try
		        {
		            //Write the workbook in file system
		        	File file = new File("src/main/resources/Students.xlsx");
		            FileOutputStream out = new FileOutputStream(file);
		            workbook.write(out);
		            out.close();
		            System.out.println("Students.xlsx written successfully on disk.");
		        } 
		        catch (Exception e) 
		        {
		            e.printStackTrace();
		        }
	}


}
