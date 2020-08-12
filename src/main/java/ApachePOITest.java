
import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlString;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class ApachePOITest {
    public static void main(String[] args) {
        Map<String, String> keywords = new HashMap<String, String>();
        keywords.put("#ISIM#", "BUSE");
        keywords.put("#TARIH#", "08/07/2020");
        keywords.put("#IZIN_SEBEBI#", "YILLIK IZIN");
        keywords.put("Country", "Ülke");
        keywords.put("Age", "26");
        keywords.put("Female", "Kadın");
        keywords.put("Template", "Taslak");
        keywords.put("document", "BUSE");
        keywords.put("or", "ODACI");
        keywords.put("IEEE", "EHIEHI");
        readAndReplaceWord("src\\main\\resources\\file-sample-doc.doc", keywords);
    }

    public static void readAndReplaceWord(String inputFile, Map<String, String> keywords) {
        System.out.println(keywords);
        try {
            File file = new File(inputFile);
            ApachePOITest apachePOITest = new ApachePOITest();
            String extension = apachePOITest.getFileExtension(file);
            if (!file.exists()) {
                throw new IOException("File does not exist!");
            }
            FileInputStream is = new FileInputStream(file);
            for (String i : keywords.keySet()) {
                if (extension.equals(".docx")) {
                    XWPFDocument doc = new XWPFDocument(is);
                    for (XWPFParagraph p : doc.getParagraphs()) {
                        List<XWPFRun> runs = p.getRuns();
                        if (runs != null) {
                            for (XWPFRun r : runs) {
                                String text = r.getText(0);
                                if (text != null && !"".equals(text.trim())) {
                                    if (text.contains(i)) {
                                        text = text.replaceAll(i, keywords.get(i));
                                        System.out.println(text);
                                        r.setText(text, 0);
                                    }
                                }
                            }
                        }
                    }
                    for (XWPFTable tbl : doc.getTables()) {
                        for (XWPFTableRow row : tbl.getRows()) {
                            for (XWPFTableCell cell : row.getTableCells()) {
                                for (XWPFParagraph p : cell.getParagraphs()) {
                                    for (XWPFRun r : p.getRuns()) {
                                        String text = r.getText(0);
                                        if (text != null && !"".equals(text.trim())) {
                                            if (text.contains(i)) {
                                                text = text.replaceAll(i, keywords.get(i));
                                                System.out.println(text);
                                                r.setText(text, 0);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    doc.write(new FileOutputStream(inputFile));
                } else if (extension.equals(".doc")) {
                    POIFSFileSystem fs = null;
                    fs = new POIFSFileSystem(new FileInputStream(inputFile));
                    HWPFDocument doc = new HWPFDocument(fs);
                    Range range = doc.getRange();
                    System.out.println(range);
                    for (int y = 0; y < range.numSections(); ++y) {
                        Section s = range.getSection(y);
                        for (int x = 0; x < s.numParagraphs(); x++) {
                            Paragraph p = s.getParagraph(x);
                            for (int z = 0; z < p.numCharacterRuns(); z++) {
                                CharacterRun run = p.getCharacterRun(z);
                                String text = run.text();
                                if (text.contains(i)) {
                                    run.replaceText(i, keywords.get(i));
                                }
                            }
                        }
                    }
                    doc.write(new FileOutputStream(inputFile));
                } else if (extension.equals(".xlsx")) {
                    XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(inputFile));
                    DataFormatter formatter = new DataFormatter();
                    for (XSSFSheet sheet : xssfWorkbook) {
                        for (Row row : sheet) {
                            for (Cell cell : row) {
                                if (formatter.formatCellValue(cell).equals(i)) {
                                    cell.setCellValue(keywords.get(i));
                                }
                            }
                        }
                    }
                    xssfWorkbook.write(new FileOutputStream(inputFile));
                } else if (extension.equals(".xls")) {
                    HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(inputFile));
                    HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            if (cell.toString().equals(i)) {
                                cell.setCellValue(keywords.get(i));
                            }
                        }
                    }
                    hssfWorkbook.write(new FileOutputStream(inputFile));
                } else if (extension.equals(".pptx")) {
                    XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(inputFile));
                    for (XSLFSlide slide : slideShow.getSlides()) {
                        CTSlide ctSlide = slide.getXmlObject();
                        XmlObject[] allText = ctSlide.selectPath(
                                "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
                                        ".//a:t");
                        for (int o = 0; o < allText.length; o++) {
                            if (allText[o] instanceof XmlString) {
                                XmlString xmlString = (XmlString) allText[o];
                                String text = xmlString.getStringValue();
                                if (text.contains(i)) {
                                    String newText = text.replaceAll(i, keywords.get(i));
                                    xmlString.setStringValue(newText);
                                }
                            }
                        }
                    }
                    slideShow.write(new FileOutputStream(inputFile));
                } else if (extension.equals(".ppt")) {
                    HSLFSlideShow slideShow = new HSLFSlideShow(new FileInputStream(inputFile));


                    slideShow.write(new FileOutputStream(inputFile));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String getFileExtension(File file) {
        String name = file.getName();
        int lastIndexOf = name.lastIndexOf(".");
        if (lastIndexOf == -1) {
            return "";
        }
        return name.substring(lastIndexOf);
    }
}
