import java.io.File;  
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
public class test{
    
    public static boolean isPalindrome(String str){  
        StringBuilder sb=new StringBuilder(str);  
        sb.reverse();  
        String rev=sb.toString();  
        if(str.equalsIgnoreCase(rev)){  
            return true;  
        }else{  
            return false;  
        }  
    } 
	
    public static void main(String[] args){
        try{
            File file = new File("StudentsEnrolledByWeek.xlsx");   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            //creating Workbook instance that refers to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            int rowCounter = 0;
            int palindromeNameCounter=0;
            List<String> palindromeNames = new ArrayList<String>();
            while (itr.hasNext()){
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()){
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                        System.out.print(cell.getStringCellValue() + "\t");
                        if(cell.getColumnIndex()==0) {
                        	if(isPalindrome(cell.getStringCellValue())) {
                        		palindromeNames.add(cell.getStringCellValue());
                        	}
                        }
                        break;
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type
                        System.out.print(cell.getNumericCellValue() + "\t");  //Operation
                        break;
                        default:
                    }
                } 
                System.out.println(""); 
                rowCounter++;//Operation
            }
            System.out.println(rowCounter);
            System.out.println(palindromeNames);
        }
        catch(Exception e){
            e.printStackTrace();
        }  

    }
    
}