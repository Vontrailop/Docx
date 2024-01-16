package docx.prueba;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.bind.JAXBElement;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtPr;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

public class FindWordAndReplace {

    private String toFind;
    private boolean startAgain;

    public FindWordAndReplace(String toFind) {

        this.toFind = toFind;
    }

    public int wordOccurances(File file) throws Docx4JException {

        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(file);
        return findWord(wmlPackage, toFind);
    }

    private int findWord(WordprocessingMLPackage doc, String toFind) {

        HashMap<ContentAccessor, List<Text>> caMap = new HashMap<ContentAccessor, List<Text>>();

        List<Object> bodyChildren = doc.getMainDocumentPart().getContent();

        for (Object child : bodyChildren) {
            if (child instanceof JAXBElement)
                child = ((JAXBElement<?>) child).getValue();

            if (child instanceof SdtBlock) {
                SdtBlock stdBlock = (SdtBlock) child;
                if (!checkIfInclude(stdBlock.getSdtPr())) {
                    do {
                        startAgain = false;
                        for (Object o : stdBlock.getSdtContent().getContent()) {

                            if (o instanceof JAXBElement)
                                o = ((JAXBElement<?>) o).getValue();
                            if (o instanceof SdtBlock) {
                                stdBlock = (SdtBlock) o;
                                startAgain = true;
                                break;
                            } else if (o instanceof ContentAccessor) {

                                ContentAccessor caElement = (ContentAccessor) o;
                                if (o instanceof P) {
                                    caMap.put(caElement, getAllTextfromContenAccessor(caElement, caMap));
                                } else {
                                    getAllTextfromContenAccessor(caElement, caMap);
                                }
                            }
                        }
                    } while (startAgain);
                }
            } else if (child instanceof ContentAccessor) {

                ContentAccessor caElement = (ContentAccessor) child;
                if (child instanceof P) {
                    caMap.put(caElement, getAllTextfromContenAccessor(caElement, caMap));
                } else {

                    getAllTextfromContenAccessor(caElement, caMap);
                }
            }
        }

        // i've the map paragraph -- textList

        int wordOcc = 0;
        for (ContentAccessor ca : caMap.keySet()) {
            if (!caMap.get(ca).isEmpty()) {
                StringBuilder builder = new StringBuilder();
                for (Text text : caMap.get(ca)) {
                    builder.append(text.getValue());
                }

                wordOcc += numOfOccourences(builder, toFind);
            }
        }

        return wordOcc;
    }

    private int numOfOccourences(StringBuilder builder, String toFind) {
        String[][] tasks = {
                { "^t", "\t" },
                { "^=", "\u2013" },
                { "^+", "\u2014" },
                { "^s", "\u00A0" },
                { "^?", "." },
                { "^#", "\\d" },
                { "^$", "\\p{L}" }
        };

        for (String[] replacement : tasks)
            toFind = toFind.replace(replacement[0], replacement[1]);

        Pattern p = Pattern.compile(toFind, Pattern.CASE_INSENSITIVE);
        Matcher m = p.matcher(builder.toString());

        int count = 0;
        while (m.find()) {
            count += 1;
        }
        return count;
    }

    /*
     * check if it is a include object
     * 
     */
    private boolean checkIfInclude(SdtPr sdtPr) {
        for (Object child : sdtPr.getRPrOrAliasOrLock()) {
            if (child instanceof JAXBElement)
                child = ((JAXBElement<?>) child).getValue();

            if (child instanceof SdtPr.Alias) {
                SdtPr.Alias alias = (SdtPr.Alias) child;
                if (alias.getVal().contains(("Include :"))) {
                    return true;
                } else
                    return false;
            }
        }
        return false;
    }

    private List<Text> getAllTextfromContenAccessor(ContentAccessor ca, HashMap<ContentAccessor, List<Text>> caMap) {

        List<Text> textList = new ArrayList<Text>();
        List<Object> children = ca.getContent();
        for (Object child : children) {
            if (child instanceof JAXBElement)
                child = ((JAXBElement<?>) child).getValue();
            if (child instanceof Text) {
                Text text = (Text) child;
                textList.add(text);
            } else if (child instanceof R) {

                R run = (R) child;
                for (Object o : run.getContent()) {

                    if (o instanceof JAXBElement)
                        o = ((JAXBElement<?>) o).getValue();

                    if (o instanceof R.Tab) {
                        Text text = new Text();
                        text.setValue("\t");

                        textList.add(text);
                    }
                    if (o instanceof R.SoftHyphen) {
                        Text text = new Text();
                        text.setValue("\u00AD");
                        textList.add(text);
                    }

                    if (o instanceof Text) {
                        textList.add((Text) o);
                    }
                }
            } else if (child instanceof ContentAccessor) {
                ContentAccessor caElement = (ContentAccessor) child;
                if (child instanceof P) {
                    caMap.put(caElement, getAllTextfromContenAccessor(caElement, caMap));
                } else {
                    getAllTextfromContenAccessor(caElement, caMap);
                }
            } else if (child instanceof SdtRun) {
                SdtRun sdtRun = (SdtRun) child;
                getAllTextFromSdtRun(sdtRun, textList, caMap);
            }
        }
        return textList;
    }

    public List<Text> getAllTextFromSdtRun(SdtRun sdtRun, List<Text> textList,
            HashMap<ContentAccessor, List<Text>> caMap) {

        if (!checkIfInclude(sdtRun.getSdtPr())) {
            for (Object o : sdtRun.getSdtContent().getContent()) {

                if (o instanceof JAXBElement)
                    o = ((JAXBElement<?>) o).getValue();

                if (o instanceof R) {

                    R run = (R) o;
                    for (Object ob : run.getContent()) {
                        if (ob instanceof JAXBElement)
                            ob = ((JAXBElement<?>) ob).getValue();

                        if (o instanceof R.Tab) {
                            Text text = new Text();
                            text.setValue("\t");
                            textList.add(text);
                        }
                        if (o instanceof R.SoftHyphen) {
                            Text text = new Text();
                            text.setValue("\u00AD");
                            textList.add(text);
                        }
                        if (ob instanceof Text) {
                            textList.add((Text) ob);
                        }
                    }
                } else if (o instanceof ContentAccessor) {

                    ContentAccessor caElement = (ContentAccessor) o;
                    if (o instanceof P) {
                        caMap.put(caElement, getAllTextfromContenAccessor(caElement, caMap));
                    } else {
                        textList.addAll(getAllTextfromContenAccessor(caElement, caMap));
                    }
                }
            }
        }
        return textList;
    }

    private void replaceWord(WordprocessingMLPackage doc, String toFind, String replacement) {
        HashMap<ContentAccessor, List<Text>> caMap = new HashMap<>();

        List<Object> bodyChildren = doc.getMainDocumentPart().getContent();

        for (Object child : bodyChildren) {
            if (child instanceof JAXBElement)
                child = ((JAXBElement<?>) child).getValue();

            if (child instanceof SdtBlock) {
                SdtBlock stdBlock = (SdtBlock) child;
                if (!checkIfInclude(stdBlock.getSdtPr())) {
                    do {
                        startAgain = false;
                        for (Object o : stdBlock.getSdtContent().getContent()) {
                            if (o instanceof JAXBElement)
                                o = ((JAXBElement<?>) o).getValue();
                            if (o instanceof SdtBlock) {
                                stdBlock = (SdtBlock) o;
                                startAgain = true;
                                break;
                            } else if (o instanceof ContentAccessor) {
                                ContentAccessor caElement = (ContentAccessor) o;
                                if (o instanceof P) {
                                    caMap.put(caElement, getAllTextfromContenAccessor(caElement, caMap));
                                } else {
                                    getAllTextfromContenAccessor(caElement, caMap);
                                }
                            }
                        }
                    } while (startAgain);
                }
            } else if (child instanceof ContentAccessor) {
                ContentAccessor caElement = (ContentAccessor) child;
                if (child instanceof P) {
                    caMap.put(caElement, getAllTextfromContenAccessor(caElement, caMap));
                } else {
                    getAllTextfromContenAccessor(caElement, caMap);
                }
            }
        }

        // Iterate through the map and replace the word
        for (ContentAccessor ca : caMap.keySet()) {
            if (!caMap.get(ca).isEmpty()) {
                for (Text text : caMap.get(ca)) {
                    String oldValue = text.getValue();
                    String newValue = oldValue.replaceAll(Pattern.quote(toFind), replacement);
                    text.setValue(newValue);
                }
            }
        }
    }

    private void findAndReplaceInHeader(WordprocessingMLPackage doc, String toFind, String replacement) {
        MainDocumentPart mainDocumentPart = doc.getMainDocumentPart();
        RelationshipsPart relsPart = mainDocumentPart.getRelationshipsPart();
        List<Relationship> headerRels = relsPart.getRelationshipsByType(Namespaces.HEADER);

        for (Relationship rel : headerRels) {
            HeaderPart headerPart = (HeaderPart) relsPart.getPart(rel);

            // Imprimir información sobre la cabecera (opcional, para depuración)
            System.out.println("Cabecera encontrada.");

            // Imprimir información sobre el contenido de la cabecera (opcional, para
            // depuración)
            System.out.println(
                    "Contenido de la cabecera: " + XmlUtils.marshaltoString(headerPart.getJaxbElement(), true, true));
            String texto = XmlUtils.marshaltoString(headerPart.getJaxbElement(), true, true);
            // headerPart.setJaxbElement(null);

            // Realizar búsqueda y reemplazo en el contenido de la cabecera
            findAndReplaceInContentAccessor(headerPart.getContent(), toFind, replacement);
        }
    }

    private void findAndReplaceInContentAccessor(List<Object> content, String toFind, String replacement) {
        for (Object child : content) {
            if (child instanceof JAXBElement) {
                child = ((JAXBElement<?>) child).getValue();
            }

            if (child instanceof SdtBlock) {
                // ... (resto del código para manejar SdtBlock)
            } else if (child instanceof ContentAccessor) {
                ContentAccessor caElement = (ContentAccessor) child;
                if (child instanceof P) {
                    findAndReplaceInTextList(((P) child).getContent(), toFind, replacement);
                } else if (child instanceof Tbl) {
                    findAndReplaceInTable((Tbl) child, toFind, replacement);
                } else {
                    findAndReplaceInContentAccessor(caElement.getContent(), toFind, replacement);
                }
            }
        }
    }

    private void findAndReplaceInTables(List<Object> content, String toFind, String replacement) {
        for (Object child : content) {
            if (child instanceof JAXBElement) {
                child = ((JAXBElement<?>) child).getValue();
            }

            if (child instanceof Tbl) {
                findAndReplaceInTable((Tbl) child, toFind, replacement);
            } else if (child instanceof ContentAccessor) {
                ContentAccessor caElement = (ContentAccessor) child;
                findAndReplaceInTables(caElement.getContent(), toFind, replacement);
            }
        }
    }

    private void findAndReplaceInTable(Tbl table, String toFind, String replacement) {
        List<Object> rows = table.getContent();
        for (Object row : rows) {
            if (row instanceof Tr) {
                List<Object> cells = ((Tr) row).getContent();
                for (Object cell : cells) {
                    if (cell instanceof Tc) {
                        List<Object> cellContent = ((Tc) cell).getContent();
                        findAndReplaceInContentAccessor(cellContent, toFind, replacement);
                    }
                }
            }
        }
    }

    private void findAndReplaceInTextList(List<Object> textList, String toFind, String replacement) {
        for (Object textObject : textList) {
            if (textObject instanceof Text) {
                Text text = (Text) textObject;
                String oldValue = text.getValue();
                String newValue = oldValue.replaceAll(Pattern.quote(toFind), replacement);
                text.setValue(newValue);
            }
        }
    }

    public static void main(String[] args) {
        String origin = System.getProperty("user.home") + "/1. OFICIO DE COMISIÓN.docx";
        String destiny = System.getProperty("user.home") + "/1.docx";
        String buscado = "folioInspeccion", valorReemplazo = "155608";

        FindWordAndReplace th = new FindWordAndReplace("fundamento");
        try {
            WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(new java.io.File(origin));
            th.replaceWord(wmlPackage, buscado, valorReemplazo);

            th.findAndReplaceInHeader(wmlPackage, buscado, valorReemplazo);

            wmlPackage.save(new java.io.File(destiny));

        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }
}