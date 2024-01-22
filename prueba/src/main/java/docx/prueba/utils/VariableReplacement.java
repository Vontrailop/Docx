package docx.prueba.utils;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.HashMap;

public class VariableReplacement {

    public void variableReplacement(File referenceFile, String[] headerVariables, HashMap<String, String> bodyVariables) {
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(referenceFile);
            VariablePrepare.prepare(wordMLPackage);

            // REEMPLAZO DE VARIABLES EN EL CUERPO DEL DOCUMENTO
            wordMLPackage.getMainDocumentPart().variableReplace(bodyVariables);

            // 1.- Obtener la sección del encabezado
            SectionWrapper section = wordMLPackage.getDocumentModel().getSections().get(0);
            if(section.getHeaderFooterPolicy() != null) {
                // 2.- Pasar la referencia del encabezado para manipular
                HeaderPart header = section.getHeaderFooterPolicy().getHeader(0);

                // 3.- Iterar los elementos del encabezado
                for(Object headerObject: header.getContent()) {
                    if(headerObject instanceof JAXBElement) {
                        if( ((JAXBElement) headerObject).getValue() instanceof Tbl ) {
                            Tbl table = (Tbl) ((JAXBElement) headerObject).getValue();

                            // PUNTO DE ITERACIÓN
                            // 4.- Obtener las filas con la variable a reemplazar
                            for(int i = 0; i < table.getContent().size(); i++) {
                                Object row = table.getContent().get(i);
                                if(row instanceof Tr) {
                                    // 5.- Obtener la celda con la variable
                                    Object cell = ((Tr) row).getContent().get(i < 5 ? 2 : 0);
                                    Tc data = (Tc) ((JAXBElement) cell).getValue();

                                    // 6.- Obtener el formato del párrafo
                                    P paragraph = (P) data.getContent().get(0);

                                    // 7.- Reemplazo de la variable
                                    replacement(paragraph, headerVariables[i]);
                                }
                            }
                        }
                    }
                }
            }

            wordMLPackage.save(new File("/Users/alcoker/Desktop/Docx/prueba/src/main/resources/templates/commission/AGv1.docx"));
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private void replacement(P paragraph, String newText) {
        // 1.- Obtener el formato del texto
        R run = (R) paragraph.getContent().get(0);

        // 2.- Limpiar el contenido del párrafo y del run
        paragraph.getContent().clear();
        run.getContent().clear();

        // 3.- Asignar el run al párrafo
        paragraph.getContent().add(run);

        // 4.- Asignar el texto al run
        Text text = new Text();
        text.setValue(newText);
        run.getContent().add(text);
    }
}
