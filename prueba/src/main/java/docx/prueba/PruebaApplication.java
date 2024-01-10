package docx.prueba;

import java.io.File;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PruebaApplication {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(PruebaApplication.class, args);

		// WordprocessingMLPackage wordmlp = WordprocessingMLPackage.createPackage();

		// wordmlp.save(new File("Helloworld.docx"));

		File doc = new File("C:/Desarrollo/Docx/prueba/FileFormat.docx");
		System.out.println(doc);
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);

		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

		String xpath = "//w:r[w:t[contains(text(),'auditFindings')]]";

		List<Object> list = documentPart.getJAXBNodesViaXPath(xpath, true);


		// java.util.HashMap mappings = new java.util.HashMap();
		// VariablePrepare.prepare(wordMLPackage);// see notes
		// mappings.put("myField", "foo");
		// wordMLPackage.getMainDocumentPart().variableReplace(mappings);

		wordMLPackage.save(new File("C:/Desarrollo/Docx/prueba/FileFormat.docx"));
	}


	











}
