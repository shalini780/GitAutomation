import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class datadriventest {
	
	public ArrayList<String> getData(String testcasename) throws IOException
	{
		//Identify  Testcase column by scanning entire 1st row
		//once coloumn is identified then scan entire  testcase column to identify purchase testcase row
		//after your grab purchase testcase row = pull all the data of that row and feed into test
		
		
		//create object for class xssfworkbook
		//workbook knows the path of excel.so for that we have define the path
		//fileInputstream arugment and pass inside the assfworkbook
		//this actually opens a channel to read your data
		//C:\Users\Shalu didi\Documents\Shalu Docs\Selenium docs
		
		//so now fis onject has access to goto the excel file so pass in the xssfworkbook.so xssfwork has an access to or move into the excel
		
		
		ArrayList<String> a= new ArrayList<String>();
		
		FileInputStream fis=new FileInputStream("C://Users//Shalu didi//Documents//Shalu Docs//Selenium docs//demodata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		//now there are multiple sheets in excel but u want to go in 1 sheet only
		//get count of sheet and compare the sheet name to the sheet which need to check
		
		int sheets = workbook.getNumberOfSheets();
		
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1"))
	{
			
			XSSFSheet sheet = workbook.getSheetAt(i);
			//Identify  Testcase column by scanning entire 1st row
			//sheet has the entire row . so we use the iterator method here
			//sheets is a collecetion of rows
			Iterator<Row> rows =sheet.iterator(); //this rows has an ability to iterate in each row in excel
			
			Row firstrow = rows.next();//using this it move in next row
			
			//now i have to read each cell ina row to know my desired column
			Iterator<Cell> ce=	firstrow.cellIterator(); //rows is a collection of cells
			
			int k=0;
			int column = 0;
			
			while(ce.hasNext())
			{
				Cell value= ce.next();
				if(value.getStringCellValue().equalsIgnoreCase("Testcase"))
				{
					//desired column
					//to get the correct column we need to get the index value
					column=k;
					
				}
				k++;
			}
			//now u have to compare each cell with your desired cell'testcase'

			System.out.println(column);
			
			//once coloumn is identified then scan entire  testcase column to identify purchase testcase row
			
			while(rows.hasNext())
			{
				
			Row r = rows.next();
			if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcasename))
			{
				//after your grab purchase testcase row = pull all the data of that row and feed into test

				Iterator<Cell> cv= r.cellIterator();
				
				while(cv.hasNext())
				{
					Cell c= cv.next();
					if(c.getCellTypeEnum()==CellType.STRING)
					{
					a.add(c.getStringCellValue());
					}
					else
					{
						a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
					}
				}
				
			}
			
			}

			
			
	}
		}
		return a;
		
		
		
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		


	}

}
