package edu.asu.main;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Convert {

	public static void main(String[] args) {
		InputStream inp;
		PrintWriter out;
		int count = 0;
		
		try {
			inp = new FileInputStream("CourseTitleUpdate.xlsx");
			out = new PrintWriter("C:/temp/CourseTitleUpdateAddendum.sql");
			
			out.println("SET DEFINE OFF\n/");
			
			try {
				XSSFWorkbook wb = new XSSFWorkbook(inp);
				Sheet sheet = wb.getSheetAt(2);
				Row row;
				String extOrgId;
				String year;
				String subj;
				String nbr;
				String pbTitle;
				String cegTitle;
				
				Iterator<Row> iter = sheet.rowIterator();
				
				while (iter.hasNext()) {
					row = iter.next();
					if (row.getCell(0).getCellType() == 0) {
						extOrgId = Double.toString(row.getCell(0).getNumericCellValue());
					} else {
						extOrgId = row.getCell(0).getStringCellValue();
					}
					if (row.getCell(1).getCellType() == 0) {
						year = Double.toString(row.getCell(1).getNumericCellValue());
					} else {
						year = row.getCell(1).getStringCellValue();
					}
					subj = row.getCell(2).getStringCellValue();
					if (row.getCell(3).getCellType() == 0) {
						nbr = Double.toString(row.getCell(3).getNumericCellValue());
					} else {
						nbr = row.getCell(3).getStringCellValue();
					}
					pbTitle = row.getCell(4).getStringCellValue().replace("'", "''");
					cegTitle = row.getCell(5).getStringCellValue().trim().replace("'", "''");
					
					System.out.println("Parsing " + extOrgId + " " + year + " " + subj + " " + nbr);
					
					out.println("UPDATE PATHWAY.COURSES");
					out.println("SET TITLE = '"	+ cegTitle + "' ");
					out.println("WHERE ID IN (SELECT C.ID FROM PATHWAY.COURSES C");
					out.println("JOIN PATHWAY.COURSE_GROUPINGS CG ON C.COURSE_GROUPING_ID = CG.ID");
					out.println("JOIN PATHWAY.REQUIREMENTS R ON CG.REQUIREMENT_ID = R.ID");
					out.println("JOIN PATHWAY.PATHWAYS P ON R.PATHWAY_ID = P.ID");
					out.println("WHERE P.EXT_ORG_ID = '" + extOrgId + "' AND P.YEAR = '" + year + "'");
					out.println("AND C.COURSE_SUBJECT = '" + subj + "' AND C.COURSE_NUMBER = '" + nbr + "'");
					out.println("AND C.TITLE = '" + pbTitle + "')");
					out.println("/");
					
					count++;
					
					if (count % 10 == 0) {
						out.println("COMMIT\n/");
					}
				}
				
				out.println("COMMIT\n/");
				
				out.close();
				wb.close();
				
				System.out.println("Done writing " + count + " entries");

			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} 
			
			catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

}
