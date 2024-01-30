package docx.prueba;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;

import java.io.File;
import java.util.List;

public class ReplaceByTable {

    public static void main(String[] args) {
        try {
            // Cargar el documento
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File("C:/Desarrollo/Docx/prueba/Prueba2.docx"));
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            Document document = mainDocumentPart.getJaxbElement();

            // Buscar y reemplazar el marcador con la tabla
            String marker = "${INSERTAR_TABLA_AQUI}";
            Tbl table = createSampleTable(); // Método para crear la tabla que deseas insertar
            findAndReplaceMarker(document, marker, table);

            // Guardar los cambios
            wordMLPackage.save(new File("C:/Desarrollo/Docx/prueba/Resultado.docx"));
            System.out.println("Tabla insertada en la ubicación del marcador.");
        } catch (Exception e) {
            System.err.println("Error al insertar la tabla: " + e.getMessage());
        }
    }

    // Método para crear una tabla de ejemplo (puedes personalizarla según tus necesidades)
    private static Tbl createSampleTable() {
        Tbl table = new Tbl();
        for (int i = 0; i < 3; i++) {
            Tr row = new Tr();
            for (int j = 0; j < 3; j++) {
                Tc cell = new Tc();
                cell.getContent().add("Row " + (i + 1) + ", Col " + (j + 1));
                row.getContent().add(cell);
            }
            table.getContent().add(row);
        }
        return table;
    }

    // Método para buscar y reemplazar el marcador con la tabla
    private static void findAndReplaceMarker(Document document, String marker, Tbl table) {
        List<Object> paragraphs = getAllParagraphs(document.getBody());
        for (Object obj : paragraphs) {
            if (obj instanceof P) {
                P p = (P) obj;
                String text = p.toString();
                if (text.contains(marker)) {
                    // Reemplazar el marcador con la tabla
                    Body body = (Body) p.getParent();
                    int index = body.getContent().indexOf(p);
                    body.getContent().remove(index);
                    body.getContent().add(index, table);
                    break;
                }
            }
        }
    }

    // Método para obtener todos los párrafos del documento
    private static List<Object> getAllParagraphs(Body body) {
        return body.getContent();
    }
}