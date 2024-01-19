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

		// Bloque para reemplazar variables en el encabezado ---------------------------------------------------------------------------------
		// 1.- Saca el encabezado del documento: en este caso saca el encabezado de la página 1
		SectionWrapper section = wordMLPackage.getDocumentModel().getSections().get(0);
		if (section.getHeaderFooterPolicy() != null) {

			// 2.- Se pasa la referencia del encabezado a un HeaderPart para manipularlo facilmente
			HeaderPart header = section.getHeaderFooterPolicy().getHeader(0);

			// 3.- Itera los objetos del encabezado
			for (Object obj : header.getContent()) {

				// filtro para separar el P de la tabla del encabezado
				if (obj instanceof JAXBElement) {

					// filtro para traer la tabla del encabezado
					if(((JAXBElement<?>) obj).getValue() instanceof Tbl) {
						Tbl table = (Tbl) ((JAXBElement<?>) obj).getValue();

						// PUNTO ITERABLE - A PARTIR DE AQUÍ SE PUEDE MODIFICAR PARA CREAR UN REEMPLAZO DINÁMICO
						// 4.- Se obtiene la fila con la variable a reemplazar (se puede iterar)
						// Nota: esta fila corresponde a las celdas con la imagen, asunto y ${folioInspeccion}
						Object row = table.getContent().get(3);
						if(row instanceof Tr) {

							// 5.- Se obtiene el la celda contenedora de la variable
							// (no se puede iterar: el index 2 corresponde a la 3ra columna de la tabla)
							Object cell = ((Tr) row).getContent().get(2);
							Tc data = (Tc) ((JAXBElement<?>) cell).getValue();

							// 6.- Se saca el parrafo de la celda para conservar el formato del parrafo
							P paragraph = (P) data.getContent().get(0);

							// Aquí puede estar lo de iterar los Run's dentro del parrafo para sacarles
							// los Text y de esta manera ir concatenando los textos para validar el valor
							// de la variable con el key del HashMap

							// 7.- Se hace el reemplazo de la variable
							setTextInParagraph(paragraph, "Nuevo texto");
						}
					}
				}
			}
		}
		// Bloque para reemplazar variables en el encabezado ---------------------------------------------------------------------------------

		wordMLPackage.save(new File("/Users/alcoker/Desktop/Docx/prueba/Resultado.docx"));
		//wordMLPackage.save(new File("C:/Desarrollo/Docx/prueba/Resultado.docx"));
	}

	// MÉTODO PARA REMPLAZAR LA VARIABLE CONSERVANDO LAS PROPIEDADES Y FORMATO
	private static void setTextInParagraph(P paragraph, String newText) {

		// 1.- Sacar al menos un Run para conservar el formato del texto
		R run = (R) paragraph.getContent().get(0);

		// 2.- Limpiar el contenido del párrafo y del run
		paragraph.getContent().clear();
		run.getContent().clear();

		// 3.- Asignar el run al parrafo
		paragraph.getContent().add(run);

		// 4.- Agregar el texto al run
		Text newTextObj = new Text();
		newTextObj.setValue(newText);
		run.getContent().add(newTextObj);;
	}
}