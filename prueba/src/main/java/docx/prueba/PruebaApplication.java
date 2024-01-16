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
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.P;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Text;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PruebaApplication {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(PruebaApplication.class, args);

		File doc = new File("C:/Desarrollo/Docx/prueba/1. OFICIO DE COMISIÓN.docx");
		//System.out.println(doc);
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);

		java.util.HashMap mappings = new java.util.HashMap();
		VariablePrepare.prepare(wordMLPackage);// see notes
		mappings.put("folioInspeccion", "DW56165");
		mappings.put("dia", "15");
		mappings.put("anoLetra", "Dos mil veinticuatro");
		// mappings.put("myField", "486qw4d68qw");

		wordMLPackage.getMainDocumentPart().variableReplace(mappings);

		

		wordMLPackage.save(new File("C:/Desarrollo/Docx/prueba/Resultado.docx"));
	}


	
		// for (SectionWrapper section : wordMLPackage.getDocumentModel().getSections()) {
		// 	if (section.getHeaderFooterPolicy() != null) {
		// 		//System.out.println(section.getHeaderFooterPolicy().getDefaultHeader());
		// 		//System.out.println("Contenido de la cabecera: " + section.getHeaderFooterPolicy().getDefaultHeader().getContent());

		// 		// for (Object obj : section.getHeaderFooterPolicy().getDefaultHeader().getContent()) {
		// 		// 	if (obj instanceof P) {
		// 		// 		P paragraph = (P) obj;
		// 		// 		// Reemplazar directamente el texto en el párrafo

		// 		// 		Text text = (Text) paragraph.getContent().get(0);
		// 		// 		text.setValue("NuevoValor");
		// 		// 	}
		// 		// }

		// 		String xpath = "//w:r[w:t[contains(text(),'myField')]]";
		// 		List<Object> list = section.getHeaderFooterPolicy().getDefaultHeader().getJAXBNodesViaXPath(xpath, true);


		// 		//section.getHeaderFooterPolicy().getDefaultHeader().variableReplace(mappings);
		// 		//section.getHeaderFooterPolicy().getDefaultFooter().variableReplace(mappings);
		
		// 	}
		// }
}