package parsing;

import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class handler {
	
	

	public static void main(String[] args) throws IOException {
		
		 String[] columns = {"File name","Target language", "Table name", "Transcation-id", "Source","Target","Database Id"};   
		 SXSSFWorkbook workbook = new SXSSFWorkbook();
		 
		        File dir = new File("D:\\Parse\\Newfolder\\English");
		        File[] files = dir.listFiles();
		        FileFilter fileFilter = new FileFilter() {
		           public boolean accept(File file) {
		              return file.isDirectory();
		           }
		        };
		        
		        files = dir.listFiles(fileFilter);
		        
		        for (int a = 0; a< files.length; a++) {
		            File filename = files[a];
		            System.out.println(filename.toString());
		         
		           
		        	
					 Sheet sheet = workbook.createSheet(filename.getName().toString());
					 
					    Font headerFont = workbook.createFont();
				        headerFont.setBold(true);
				        headerFont.setFontHeightInPoints((short) 14);
				        headerFont.setColor(IndexedColors.RED.getIndex());
				        
				        CellStyle headerCellStyle = workbook.createCellStyle();
				        headerCellStyle.setFont(headerFont);
				        
				        Row headerRow = sheet.createRow(0);
				        
				        for(int i = 0; i < columns.length; i++) {
				            Cell cell = headerRow.createCell(i);
				            cell.setCellValue(columns[i]);
				            cell.setCellStyle(headerCellStyle);
				        }
				        
		
		 try {
			 String innerFiles=filename.toString()+"\\xliff\\common_pmp";
			 File folder = new File(innerFiles);
			 File[] listOfFiles = folder.listFiles();

			 int rowNum=1;
			 
			     for (int i = 0; i < listOfFiles.length; i++) {
			     
				String filepath = innerFiles+"\\"+listOfFiles[i].getName();
				DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
				Document doc = docBuilder.parse(filepath);
				
				doc.getDocumentElement().normalize();
				
				System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
				
				
				Row row=sheet.createRow(rowNum++);
				
				
				NodeList nList1 = doc.getElementsByTagName("file");
				NodeList nList2 = doc.getElementsByTagName("trans-unit");
				
				Node nNode1=nList1.item(0);
				Element eElement1 = (Element) nNode1;
				
				String fileName =listOfFiles[i].getName().toString();
				row.createCell(0).setCellValue(fileName);
				row.createCell(1).setCellValue(eElement1.getAttribute("target-language"));
				row.createCell(2).setCellValue(eElement1.getElementsByTagName("sup:Table").item(0).getTextContent());
				
				System.out.println("Target-language : " + eElement1.getAttribute("target-language"));
				System.out.println("Table name : " + eElement1.getElementsByTagName("sup:Table").item(0).getTextContent()+"\n");
				
				for (int temp = 0; temp < nList2.getLength(); temp++) {

					Node nNode2 = nList2.item(temp);
							
					//System.out.println("\nCurrent Element : " + nNode.getNodeName());
							
					if (nNode2.getNodeType() == Node.ELEMENT_NODE) {

						Row row1=sheet.createRow(rowNum++);
						Element eElement2 = (Element) nNode2;
						
						System.out.println("Transaction id : " + eElement2.getAttribute("id"));
						row1.createCell(3).setCellValue(eElement2.getAttribute("id"));
						System.out.println("Source : " + eElement2.getElementsByTagName("source").item(0).getTextContent());
						row1.createCell(4).setCellValue(eElement2.getElementsByTagName("source").item(0).getTextContent());
						System.out.println("Target : " + eElement2.getElementsByTagName("target").item(0).getTextContent());
						row1.createCell(5).setCellValue(eElement2.getElementsByTagName("target").item(0).getTextContent());
						
						System.out.println("Database ID : " + eElement2.getElementsByTagName("sup:Reference").item(1).getTextContent()+"\n");
						row1.createCell(6).setCellValue(eElement2.getElementsByTagName("sup:Reference").item(1).getTextContent());
						
						
						
					}
				}	
				
				/*for(int x = 0; x < columns.length; x++) {
		            sheet.autoSizeColumn(x);
		        }
			*/	
				
				
				
				}
			     
			     
			    
		 }
		 catch(Exception e)
		 {
			 e.printStackTrace();
		 }
		 
		
		 }
		        
		        FileOutputStream fileOut = new FileOutputStream("details.xlsx");
		        workbook.write(fileOut);
		        fileOut.close();
				 System.out.println("Successfully Created");
				 
				 
				 
		     workbook.close();
	}

}
