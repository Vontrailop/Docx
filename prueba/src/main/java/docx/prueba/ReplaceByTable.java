package docx.prueba;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.io.File;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;

public class ReplaceByTable {
    public static void main(String[] args) {
        try {
            // Cargar el documento
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(new File("C:/Desarrollo/Docx/prueba/Prueba2.docx"));
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            Document document = mainDocumentPart.getJaxbElement();

            // Buscar y reemplazar el marcador con la tabla
            String marker = "${INSERTAR_TABLA_AQUI}";

            // Crear un mapa con los datos que deseas insertar en la tabla
            Map<String, String> datos = new LinkedHashMap<>();
            datos.put("Clave1", "Valor1");
            datos.put("Clave2", "Valor2");
            datos.put("Clave3", "Valor3");

            Tbl table = createTableFromMap(datos);

            setCellBackgroundColor(table, 0, "FF0000");

            findAndReplaceMarker(document, marker, table);

            // Guardar los cambios
            wordMLPackage.save(new File("C:/Desarrollo/Docx/prueba/Resultado.docx"));
            System.out.println("Tabla insertada en la ubicación del marcador.");
        } catch (Exception e) {
            System.err.println("Error al insertar la tabla: " + e.getMessage());
        }
    }

    private static Tbl createTableFromMap(Map<String, String> datos) {
        Tbl table = new Tbl();
        for (Map.Entry<String, String> entry : datos.entrySet()) {
            Tr row = new Tr();

            // Celda de la clave
            Tc cell1 = createTableCell(entry.getKey());

            // Celda del valor
            Tc cell2 = createTableCell(entry.getValue());

            // Agregar las celdas a la fila
            row.getContent().add(cell1);
            row.getContent().add(cell2);

            // Agregar la fila a la tabla
            table.getContent().add(row);
        }
        return table;
    }

    private static Tc createTableCell(String text) {
        Tc cell = new Tc();
        P p = new P();
        R run = new R();
        Text t = new Text();
        t.setValue(text);
        run.getContent().add(t);
        p.getContent().add(run);
        cell.getContent().add(p);
        return cell;
    }

    // Método para buscar y reemplazar el marcador con la tabla
    private static void findAndReplaceMarker(Document document, String marker, Tbl table) {
        List<Object> paragraphs = document.getBody().getContent();
        for (int i = 0; i < paragraphs.size(); i++) {
            Object obj = paragraphs.get(i);
            if (obj instanceof P) {
                P p = (P) obj;
                String text = p.toString();
                if (text.contains(marker)) {
                    // Reemplazar el marcador con la tabla
                    document.getBody().getContent().remove(i);
                    document.getBody().getContent().add(i, table);
                    break;
                }
            }
        }
    }

    private static void setCellBackgroundColor(Tbl table, int columnIndex, String colorHex) {
        List<Object> rows = table.getContent();
        for (Object row : rows) {
            if (row instanceof JAXBElement) {
                Object obj = ((JAXBElement<?>) row).getValue();
                if (obj instanceof org.docx4j.wml.Tr) {
                    org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
                    List<Object> cells = tr.getContent();
                    if (cells.size() > columnIndex && cells.get(columnIndex) instanceof JAXBElement) {
                        Object cellObj = ((JAXBElement<?>) cells.get(columnIndex)).getValue();
                        if (cellObj instanceof Tc) {
                            Tc tc = (Tc) cellObj;

                            // Crear un nuevo objeto de propiedades de celda (TcPr)
                            TcPr tcPr = Context.getWmlObjectFactory().createTcPr();

                            // Crear un nuevo objeto de propiedades de fondo de celda (CTShd)
                            org.docx4j.wml.CTShd ctShd = new org.docx4j.wml.CTShd();
                            ctShd.setFill(colorHex); // Establecer el color en formato hexadecimal
                            tcPr.setShd(ctShd);

                            // Aplicar las propiedades de la celda
                            tc.setTcPr(tcPr);
                        }
                    }
                }
            }
        }
    }
}