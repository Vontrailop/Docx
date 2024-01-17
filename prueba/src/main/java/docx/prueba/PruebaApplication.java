package docx.prueba;

import java.io.File;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.com.microsoft.schemas.office.word.x2006.wordml.CTRel;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PruebaApplication {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(PruebaApplication.class, args);

		File doc = new File("/Users/alcoker/Desktop/Docx/prueba/1. OFICIO DE COMISIÓN.docx");
		//File doc = new File("C:/Desarrollo/Docx/prueba/1. OFICIO DE COMISIÓN.docx");

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);

		java.util.HashMap mappings = new java.util.HashMap();
		VariablePrepare.prepare(wordMLPackage);// see notes
		mappings.put("folioInspeccion", "DW56165");
		mappings.put("dia", "15");
		mappings.put("anoLetra", "Dos mil veinticuatro");


		wordMLPackage.getMainDocumentPart().variableReplace(mappings);

		SectionWrapper section = wordMLPackage.getDocumentModel().getSections().get(0);
		if (section.getHeaderFooterPolicy() != null) {
			HeaderPart header = section.getHeaderFooterPolicy().getHeader(0);

			for (Object obj : header.getContent()) {
				if (obj instanceof JAXBElement) {
					if(((JAXBElement<?>) obj).getValue() instanceof Tbl) {
						Tbl table = (Tbl) ((JAXBElement<?>) obj).getValue();

						for(Object row : table.getContent()) {
							if(row instanceof Tr) {

								for(Object cell : ((Tr) row).getContent()) {
									Tc tCell = (Tc) ((JAXBElement<?>) cell).getValue();

									setTextInCell(tCell, "FolioNuevoDeLaInspección");
								}
							}
						}
					}
				}
			}

			String xpath = "w:r[w:t[contains(text(),'myField')]]";
			List<Object> list = section.getHeaderFooterPolicy().getDefaultHeader().getJAXBNodesViaXPath(xpath, true);


			section.getHeaderFooterPolicy().getDefaultHeader().variableReplace(mappings);
			section.getHeaderFooterPolicy().getDefaultFooter().variableReplace(mappings);

		}

		wordMLPackage.save(new File("/Users/alcoker/Desktop/Docx/prueba/Resultado.docx"));
		//wordMLPackage.save(new File("C:/Desarrollo/Docx/prueba/Resultado.docx"));
	}

//	private static String getTextInCell(Tc cell) {
//		StringBuilder text = new StringBuilder();
//		List<Object> cellContent = cell.getContent();
//		for (Object obj : cellContent) {
//			if (obj instanceof P) {
//				P paragraph = (P) obj;
//				text.append(getTextInParagraph(paragraph));
//			}
//		}
//		return text.toString();
//	}
//
//	private static String getTextInParagraph(P paragraph) {
//		StringBuilder text = new StringBuilder();
//		List<Object> texts = paragraph.getContent();
//		for (Object textObj : texts) {
//			if (textObj instanceof Text) {
//				Text t = (Text) textObj;
//				text.append(t.getValue());
//			}
//		}
//		return text.toString();
//	}

	private static void setTextInCell(Tc cell, String newText) {
		// Crear un nuevo run
		R newRun = new R();

		// Crear un nuevo párrafo con el nuevo run
		P newParagraph = new P();
		newParagraph.getContent().add(newRun);

		// Agregar el texto al nuevo run
		Text newTextObj = new Text();
		newTextObj.setValue(newText);
		newRun.getContent().add(newTextObj);

		// Agregar el nuevo párrafo al contenido de la celda
		cell.getContent().add(newParagraph);
	}
}