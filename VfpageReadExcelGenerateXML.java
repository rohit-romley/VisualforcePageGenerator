import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class VfpageReadExcelGenerateXML {
	
	public static Map makeMap(){
		Map fieldsMap = new HashMap();
		fieldsMap.put("field1","field1__c");
    fieldsMap.put("field2","field2__c");
    fieldsMap.put("field3","field3__c");
    fieldsMap.put("field4","field4__c");
		
		return fieldsMap;

	}
	public static void main(String argv[]){
		try {
			//layout.xlsx contains the basic layout of the page you'd like to generate
			String fileName = "C:\\layout.xlsx";
			String objectName = "ObjectName__c";
			Map fieldsMap = makeMap();
				
				/**XML file initialization START*/
				DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
	
				// root elements
				Document doc = docBuilder.newDocument();
				Element rootElement = doc.createElement("apex:page");
				rootElement.setAttribute("standardController", objectName);
				doc.appendChild(rootElement);
				
				//to Display Errors
				Element pageMessages = doc.createElement("apex:pageMessages");
				doc.appendChild(pageMessages);
				
				Element form =  doc.createElement("apex:form");
				rootElement.appendChild(form);
				
				
				
				Element pageBlock = null;
	            Element pageBlockSection = null;
	            Element inputField  = null;
	            Element pageBlockSectionItem = doc.createElement("apex:pageBlockSectionItem");
	            
				/**XML file initialization END*/
			
			
			
			    /**Excel initialization START*/
			    Workbook wb = WorkbookFactory.create(new File(fileName));
			    Sheet sheet = wb.getSheetAt(0);

			    //Iterate through each rows from first sheet
			    Iterator<Row> rowIterator = sheet.iterator();
			    /**Excel initialization END*/
			    
			    while(rowIterator.hasNext()) {
			        Row row = rowIterator.next();

			        //For each row, iterate through each columns
			        Iterator<Cell> cellIterator = row.cellIterator();
			        while(cellIterator.hasNext()) {

			            Cell cell = cellIterator.next();
			            
			            switch(cell.getCellType()) {
			                case Cell.CELL_TYPE_BOOLEAN:
			                    System.out.print(cell.getBooleanCellValue() + "\t\t");
			                    break;
			                case Cell.CELL_TYPE_NUMERIC:
			                    System.out.print(cell.getNumericCellValue() + "\t\t");
			                    break;
			                case Cell.CELL_TYPE_STRING:
			                	switch(cell.getStringCellValue()){
			                		
			                	case "pbTitle":
			                		pageBlock = doc.createElement("apex:pageBlock");
			                		
			                		if(cellIterator.hasNext()){
			                			cell = cellIterator.next();
			                			pageBlock.setAttribute("title", cell.getStringCellValue());
			                			pageBlock.setAttribute("mode", "edit");
			                			form.appendChild(pageBlock);
			                			System.out.print(cell.getStringCellValue() + "\t\t\n");
			                		}
			                		Element pageBlockButtons =  doc.createElement("apex:pageBlockButtons");
			                		pageBlock.appendChild(pageBlockButtons);
			                		Element commandButton =  doc.createElement("apex:commandButton");
			                		commandButton.setAttribute("action", "{!save}");
			                		commandButton.setAttribute("value", "Save");
			                		pageBlockButtons.appendChild(commandButton);
			                		
			                		commandButton =  doc.createElement("apex:commandButton");
			                		commandButton.setAttribute("action", "{!cancel}");
			                		commandButton.setAttribute("value", "Cancel");
			                		pageBlockButtons.appendChild(commandButton);
			                		
			                		break;
			                	case "pbsTitle":
			                		pageBlockSection = doc.createElement("apex:pageBlockSection");
			                		if(cellIterator.hasNext()){
			                			cell = cellIterator.next();
			                			pageBlockSection.setAttribute("title", cell.getStringCellValue());
			                			pageBlockSection.setAttribute("columns", "2");
			                			pageBlockSection.setAttribute("showHeader", "true");
			                			pageBlock.appendChild(pageBlockSection);
			                			System.out.print(cell.getStringCellValue() + "\t\t\n");
			                		}
			                		break;
			                		
			                	case "BLANKSPACE":
			                		if(pageBlockSection!=null)
			                			pageBlockSection.appendChild(pageBlockSectionItem);
			                		else
			                			pageBlock.appendChild(pageBlockSectionItem);
			                		break;
			                		
			                	default:
			                		
			                		inputField = doc.createElement("apex:inputField");
			                		inputField.setAttribute("value", "{!"+objectName+"."+fieldsMap.get(cell.getStringCellValue())+"}");
			                		if(pageBlockSection!=null)
			                			pageBlockSection.appendChild(inputField);
			                		else
			                			pageBlock.appendChild(inputField);
			                		System.out.print(cell.getStringCellValue() + "\t\t");
			                		break;
			                	}
			                    break;
			            }
			            
			        }System.out.println("");
			        
			    }

				// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer = transformerFactory.newTransformer();
				transformer.setOutputProperty(OutputKeys.INDENT, "yes");
				transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
				transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new File("C:\\layout.xml"));

				// Output to console for testing
				// StreamResult result = new StreamResult(System.out);

				transformer.transform(source, result);

				System.out.println("File saved!");

		} catch(Exception ioe) {
		    ioe.printStackTrace();
		}
	}

}
