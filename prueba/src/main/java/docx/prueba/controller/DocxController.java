package docx.prueba.controller;

import docx.prueba.utils.VariableReplacement;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.util.HashMap;

@RestController
@CrossOrigin(origins = "*")
@RequestMapping("/api/v1/docs")
public class DocxController {

    private final VariableReplacement variableReplacement = new VariableReplacement();

    @PostMapping("")
    public String generateCommissionFile() {
        String[] headerVariables = {
                "SECRETARÍA DE DESARROLLO ECONÓMICO Y DEL TRABAJO",
                "COORDINACIÓN DEL TRABAJO Y PREVISIÓN SOCIAL",
                "DIRECCIÓN GENERAL DE INSPECCIÓN DEL TRABAJO DEL ESTADO DE MORELOS",
                "DGTI/XXXX/XXXX/AG/XXXX",
                "SE CONFIERE COMISIÓN",
                "\"2024, lema del año\""
        };

        HashMap<String, String> bodyVariables = new HashMap<>();
        bodyVariables.put("citizen", "ANGEL YAZVECK ALCOCER DURÁN");
        bodyVariables.put("fullDate", "22 de enero de dos mil veinticuatro");
        bodyVariables.put("laboralMatter", "PAGO DE AGUINALDO");
        bodyVariables.put("officeNumber", "DGTI/XXXX/XXXX/AG/XXXX");
        bodyVariables.put("businessName", "Coker Salinas S. A. de C. V.");
        bodyVariables.put("direction", "Emiliano Zapata, Morelos");
        bodyVariables.put("inspectionDate", "28/01/2024");
        bodyVariables.put("inspectionDateTime", "16:00");
        bodyVariables.put("principalName", "LIC. ZOILA MARÍA ALEJANDRA JARILLO SOTO");
        bodyVariables.put("principalJobPosition", "DIRECTORA GENERAL DE INSPECCIÓN DEL TRABAJO DEL ESTADO DE MORELOS.");

        File comissionFile = new File("/Users/alcoker/Desktop/Docx/prueba/src/main/resources/templates/commission/AG.docx");
        variableReplacement.variableReplacement(comissionFile, headerVariables, bodyVariables);
        return "Conchesumare";
    }
}
