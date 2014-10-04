package org.adrngrl.wortgen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

public class Wortgen {

	public static void main(String[] args){
		
		 
			
		try {
			FileInputStream dfis = null;
			dfis = new FileInputStream(args[0]);
			
			File output = new File(args[1]);
			FileOutputStream fos = new FileOutputStream(output);
			OutputStreamWriter osw = new OutputStreamWriter(fos, Charset.forName("Cp1250"));
			String line = null;
			
			XSSFWorkbook workbook = new XSSFWorkbook(dfis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			Iterator<Row> rowIterator = sheet.iterator();
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					/* switch(cell.getCellType())
					{
					case Cell.CELL_TYPE_STRING:
						line = cell.getStringCellValue();
						switch(cell.getColumnIndex())
						{
						case 0: osw.write(line + " ");
								break;
						case 1: osw.write(line);
								break;
						case 2: osw.write(", ");
								break;
						case 3:	osw.write(line);
								break;
						case 4: osw.write("\t");
								break;
						case 5: osw.write(line);
								break;
						}
						break;
					case Cell.CELL_TYPE_FORMULA:
						line = cell.getStringCellValue();
						switch(cell.getColumnIndex())
						{
						case 2: if(line.contains("%C")) osw.write(", ");
								break;
						case 6: osw.write(" " + line);
								break;
						}
					} */
					
					if(cell.getCellType() == Cell.CELL_TYPE_STRING || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
					{	
						line = cell.getStringCellValue();
						switch(cell.getColumnIndex())
						{
						case 0: osw.write(line + " ");			// rodzajnik
								break;
						case 1: osw.write(line);			// wlasciwy termin
								break;
						case 2: case 3: case 4: osw.write(", " + line);			// dodatkowe formy
								break;
						case 5: osw.write("\t" + line);			// tlumaczenie
								break;
						case 6: osw.write(" " + line);			// dodatkowa informacja
						}
					}				
				}
				osw.write("\r\n");
			}
			dfis.close();
			osw.close();
			
		} catch (Exception e) {
			System.out.print("Source file not found.\n");
			System.out.print("Try: wortgen.bat <source>.xlsx <destination>.txt\n");
			// e.printStackTrace();
			// test
		}
		

	}
}
