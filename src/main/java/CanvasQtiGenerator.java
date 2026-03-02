import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CanvasQtiGenerator {

    public static void main(String[] args) {
        JFileChooser chooser = new JFileChooser();
        int result = chooser.showOpenDialog(null);

        if (result == JFileChooser.APPROVE_OPTION) {
            try {
                generateQTI(chooser.getSelectedFile());
                JOptionPane.showMessageDialog(null, "Canvas QTI package created successfully.");
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null, "Error: " + e.getMessage());
            }
        }
    }

    private static void generateQTI(File xlsFile) throws Exception {

        Workbook workbook = new HSSFWorkbook(new FileInputStream(xlsFile));
        Sheet sheet = workbook.getSheetAt(0);

        String assessmentIdent = getCellValueAsString(sheet.getRow(0).getCell(1));
        String assessmentTitle = getCellValueAsString(sheet.getRow(1).getCell(1));

        if (assessmentTitle.isBlank()) {
            throw new Exception("Assessment title (cell B2) is empty.");
        }

        String safeFileName = assessmentTitle.replaceAll("[^a-zA-Z0-9\\-_]", "_");

        Path tempDir = Files.createTempDirectory("canvas_qti");

        File assessmentXml = new File(tempDir.toFile(), "assessment.xml");
        File manifestXml = new File(tempDir.toFile(), "imsmanifest.xml");

        createAssessmentXML(sheet, assessmentIdent, assessmentTitle, assessmentXml);
        createManifestXML(manifestXml);

        File zipFile = new File(safeFileName + ".zip");
        zipDirectory(tempDir.toFile(), zipFile);

        deleteDirectory(tempDir.toFile());
        workbook.close();
    }

    private static void createAssessmentXML(Sheet sheet,
                                            String assessmentIdent,
                                            String assessmentTitle,
                                            File outputFile) throws Exception {

        DocumentBuilder builder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        Document doc = builder.newDocument();

        Element questestinterop = doc.createElement("questestinterop");
        doc.appendChild(questestinterop);

        Element assessment = doc.createElement("assessment");
        assessment.setAttribute("ident", assessmentIdent);
        assessment.setAttribute("title", assessmentTitle);
        questestinterop.appendChild(assessment);

        Element section = doc.createElement("section");
        section.setAttribute("ident", "root_section");
        assessment.appendChild(section);

        for (int i = 4; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);
            if (row == null) continue;

            String questionText = getCellValueAsString(row.getCell(0));
            if (questionText.isBlank()) continue;  // Skip empty rows

            String itemIdent = getCellValueAsString(row.getCell(1));
            String itemTitle = getCellValueAsString(row.getCell(2));

            String[] options = {
                    getCellValueAsString(row.getCell(3)),
                    getCellValueAsString(row.getCell(4)),
                    getCellValueAsString(row.getCell(5)),
                    getCellValueAsString(row.getCell(6))
            };

            String correct = getCellValueAsString(row.getCell(7));

            Element item = doc.createElement("item");
            item.setAttribute("ident", itemIdent.isBlank() ? UUID.randomUUID().toString() : itemIdent);
            item.setAttribute("title", itemTitle.isBlank() ? "Question " + (i - 3) : itemTitle);
            section.appendChild(item);

            Element presentation = doc.createElement("presentation");
            item.appendChild(presentation);

            Element material = doc.createElement("material");
            Element mattext = doc.createElement("mattext");
            mattext.setAttribute("texttype", "text/plain");
            mattext.setTextContent(questionText);
            material.appendChild(mattext);
            presentation.appendChild(material);

            Element response = doc.createElement("response_lid");
            response.setAttribute("ident", "response1");
            response.setAttribute("rcardinality", "Single");
            presentation.appendChild(response);

            Element renderChoice = doc.createElement("render_choice");
            response.appendChild(renderChoice);

            String[] labels = {"A", "B", "C", "D"};

            for (int j = 0; j < labels.length; j++) {

                if (options[j].isBlank()) continue;

                Element responseLabel = doc.createElement("response_label");
                responseLabel.setAttribute("ident", labels[j]);

                Element mat = doc.createElement("material");
                Element mt = doc.createElement("mattext");
                mt.setAttribute("texttype", "text/plain");
                mt.setTextContent(options[j]);
                mat.appendChild(mt);

                responseLabel.appendChild(mat);
                renderChoice.appendChild(responseLabel);
            }

            Element resprocessing = doc.createElement("resprocessing");
            item.appendChild(resprocessing);

            Element respcondition = doc.createElement("respcondition");
            respcondition.setAttribute("continue", "No");
            resprocessing.appendChild(respcondition);

            Element conditionvar = doc.createElement("conditionvar");
            respcondition.appendChild(conditionvar);

            Element varequal = doc.createElement("varequal");
            varequal.setAttribute("respident", "response1");
            varequal.setTextContent(correct);
            conditionvar.appendChild(varequal);

            Element setvar = doc.createElement("setvar");
            setvar.setAttribute("action", "Set");
            setvar.setTextContent("100");
            respcondition.appendChild(setvar);
        }

        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.transform(new DOMSource(doc), new StreamResult(outputFile));
    }

    private static void createManifestXML(File outputFile) throws Exception {

        String identifier = UUID.randomUUID().toString();

        String manifest =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<manifest identifier=\"" + identifier + "\" " +
                "xmlns=\"http://www.imsglobal.org/xsd/imscp_v1p1\" " +
                "xmlns:imsqti=\"http://www.imsglobal.org/xsd/ims_qtiasiv1p2\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n" +
                "  <organizations/>\n" +
                "  <resources>\n" +
                "    <resource identifier=\"res1\" type=\"imsqti_xmlv1p2\" href=\"assessment.xml\">\n" +
                "      <file href=\"assessment.xml\"/>\n" +
                "    </resource>\n" +
                "  </resources>\n" +
                "</manifest>";

        Files.writeString(outputFile.toPath(), manifest);
    }

    private static void zipDirectory(File folder, File zipFile) throws Exception {
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
            for (File file : folder.listFiles()) {
                zos.putNextEntry(new ZipEntry(file.getName()));
                Files.copy(file.toPath(), zos);
                zos.closeEntry();
            }
        }
    }

    private static void deleteDirectory(File directory) {
        for (File file : directory.listFiles()) {
            file.delete();
        }
        directory.delete();
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }
}
