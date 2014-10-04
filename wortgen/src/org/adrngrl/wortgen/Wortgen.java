package org.adrngrl.wortgen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.IOException;
import java.nio.channels.ShutdownChannelGroupException;
import java.nio.charset.Charset;
import java.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

public class Wortgen {

	public static void main(String[] args){
	
	System.out.println(args[0] + "\n");
		
	List<String> InputLines = new ArrayList<String>();
					
		try {
			
			String[] tokens = args[0].split("\\.(?=[^\\.]+$)");
			String FileNameBase = tokens[0];
			
			FileInputStream dfis = null;
			dfis = new FileInputStream(args[0]);
			
			File output = new File(FileNameBase + ".txt");
			FileOutputStream fos = new FileOutputStream(output);
			OutputStreamWriter osw = new OutputStreamWriter(fos, Charset.forName("Cp1250"));
			
			
			String line = null;
			
			
			XSSFWorkbook workbook = new XSSFWorkbook(dfis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			Iterator<Row> rowIterator = sheet.iterator();
			while(rowIterator.hasNext()) {
				
				String InputLine = new String();
				
				Row row = rowIterator.next();
				
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					if(cell.getCellType() == Cell.CELL_TYPE_STRING || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
					{	
						line = cell.getStringCellValue();
						switch(cell.getColumnIndex())
						{
						case 0: osw.write(line + " ");			// rodzajnik
								InputLine = InputLine + line + " ";
								break;
						case 1: osw.write(line);			// wlasciwy termin
								InputLine = InputLine + line;
								break;
						case 2: case 3: case 4: osw.write(", " + line);			// dodatkowe formy
								InputLine = InputLine + ", " + line;
								break;
						case 5: osw.write("\t" + line);			// tlumaczenie
								InputLine = InputLine + "\t" + line;
								break;
						case 6: osw.write(" " + line);			// dodatkowa informacja
								InputLine = InputLine + " " + line;
						}
					}				
				}
				InputLines.add(InputLine);
				osw.write("\r\n");
			}
			dfis.close();
			osw.close();
			
			if(args.length > 1){
				
				int SetSize = Integer.parseInt(args[1]);
			
				Collections.shuffle(InputLines);
			
				int numOfShuffledFiles = InputLines.size() / SetSize;
				Iterator<String> ilIterator = InputLines.iterator();

				for(int i = 0; i <= numOfShuffledFiles; i++)
				{
					File soutput = new File(FileNameBase + "-" + String.format("%02d", i + 1) + "-of-" + String.format("%02d", numOfShuffledFiles + 1) + "-shuffled.txt");
					FileOutputStream sfos = new FileOutputStream(soutput);
					OutputStreamWriter sosw = new OutputStreamWriter(sfos, Charset.forName("Cp1250"));
			
							
					for(int j = 0; j < SetSize; j++)
					{
						if(!ilIterator.hasNext()) break;
						sosw.write(ilIterator.next() + "\r\n");
					}
			
					sosw.close();
				}
			}
			
		} catch (Exception e) {
			System.out.print("Source file not found.\n");
			System.out.print("Try: wortgen.bat <source>.xlsx <shuffle set size>\n");
			e.printStackTrace();
			// test
		}
		

	}
}
