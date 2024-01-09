package docx.prueba;

import java.io.File;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.soap.Text;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PruebaApplication {

	public static void main(String[] args) throws Docx4JException, JAXBException {
		SpringApplication.run(PruebaApplication.class, args);

		// WordprocessingMLPackage wordPackage =
		// WordprocessingMLPackage.createPackage();
		// MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
		// mainDocumentPart.addStyledParagraphOfText("Title", "Hello World!");
		// mainDocumentPart.addParagraphOfText("Welcome To Baeldung");
		// File exportFile = new File("welcome.docx");
		// wordPackage.save(exportFile);

		File doc = new File("Prueba1.docx");
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
		String textNodesXPath = "//w:t";
		List<Object> textNodes = mainDocumentPart
				.getJAXBNodesViaXPath(textNodesXPath, true);
		for (Object obj : textNodes) {
			Text text = (Text) ((JAXBElement) obj).getValue();
			String textValue = text.getValue();
			System.out.println(textValue);
		}

	}

}
