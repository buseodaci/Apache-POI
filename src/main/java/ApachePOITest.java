
import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.SlideShow;
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
import org.odftoolkit.odfdom.doc.OdfDocument;
import org.odftoolkit.odfdom.doc.OdfPresentationDocument;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.doc.presentation.OdfSlide;
import org.odftoolkit.odfdom.incubator.search.TextNavigation;
import org.odftoolkit.odfdom.incubator.search.TextSelection;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
import org.openxmlformats.schemas.presentationml.x2006.main.PresentationDocument;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;


public class ApachePOITest {

    public static void main(String[] args) {
        Map<String, String> keywords = new HashMap<String, String>();
        keywords.put("#ISIM#", "BUSE");
        keywords.put("#TARIH#", "08/07/2020");
        keywords.put("#IZIN_SEBEBI#", "YILLIK IZIN");
        keywords.put("Lorem", "14");
        keywords.put("ipsum", "17");

        readAndReplaceWord("src\\main\\resources\\file-sample-odp.odp", keywords);
    }

    public static void readAndReplaceWord(String inputFile, Map<String, String> keywords) {
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
                        for (XmlObject xmlObject : allText) {
                            if (xmlObject instanceof XmlString) {
                                XmlString xmlString = (XmlString) xmlObject;
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
                    SlideShow ppt = new SlideShow(new HSLFSlideShow(inputFile));
                    Slide[] slides = ppt.getSlides();
                    for (int p = 0; p < slides.length; p++) {
                        TextRun[] text = slides[p].getTextRuns();
                        for (int j = 0; j < text.length; j++) {
                            String[] words = text[j].getText().split("\\s+");
                            for (int t = 0; t < words.length; t++) {
                                if (words[t].equals(i)) {
                                    words[t].replaceAll(i, keywords.get(i));
                                }
                            }
                            System.out.println(words);
                        }
                    }
                    /*SlideShow ppt = new SlideShow(new HSLFSlideShow(inputFile));
                    Slide[] slides = ppt.getSlides();
                    if (slides == null || slides.length == 0) {
                        System.out.println(inputFile + " doesn't contains any slide.");
                        return;
                    }
                    Pattern pattern = Pattern.compile(REGEXP);
                    Matcher matcher = null;
                    StringBuilder sb = new StringBuilder();
                    for (Slide slide : slides) {
                        TextRun[] textRuns = slide.getTextRuns();
                        System.out.println("slide number:" + slide.getSlideNumber());
                        for (TextRun run : textRuns) {
                            RichTextRun[] richTextRuns = run.getRichTextRuns();
                            for (RichTextRun richTextRun : richTextRuns) {
                                sb.delete(0, sb.length());
                                sb.append(richTextRun.getText());
                                matcher = pattern.matcher(sb.toString());
                                boolean change = false;
                                while (matcher.find()) {
                                    change = true;
                                    String param = matcher.group();
                                    System.out.println("match found: " + param);
                                    //replace match in text.
                                    String replacement = keywords.get(i);
                                    System.out.println(param + " replaced by " + replacement);
                                    int start = sb.indexOf(param);
                                    int end = start + param.length();
                                    sb.replace(start, end, replacement);
                                }
                                if (change) {
                                    System.out.println("text changed");
                                    richTextRun.setText(sb.toString());
                                }
                            }
                        }
                    }
                */
                    ppt.write(new FileOutputStream(inputFile));
                    /* for (int j = 0; j < text.length; j++) {
                            String[] kelimeler = text[j].getText().split("\\s+");
                            for (int t = 0; t < text[j].getText().split("\\s+").length; t++) {
                                if (text[j].getText().split("\\s+")[t].equals(i)) {
                                    System.out.println("EHi " + text[j].getText().split("\\s+")[t]);
                                    text[j].getText().split("\\s+")[t].replaceAll(i, keywords.get(i));
                                }
                            }
                        }*/
                    /*
                        for (TextRun textRun : text) {
                            if (textRun.getText().split("\\s+").equals(i)) {
                                textRun.setText(keywords.get(i));
                            }
                        }
                    }
                    ppt.write(new FileOutputStream(inputFile));*/
                } else if (extension.equals(".odt")) {
                    OdfTextDocument odfTextDocument = (OdfTextDocument) OdfDocument.loadDocument(inputFile);
                    TextNavigation search = new TextNavigation(i, odfTextDocument);
                    int b = 0;
                    while (search.hasNext()) {
                        TextSelection item = (TextSelection) search.getCurrentItem();
                        item.replaceWith(keywords.get(i));
                        b++;
                    }
                    odfTextDocument.save(inputFile);
                } else if (extension.equals(".odp")) {
                    OdfPresentationDocument odfPresentationDocument = (OdfPresentationDocument) OdfDocument.loadDocument(inputFile);
                    System.out.println(odfPresentationDocument.getSlides());
                    odfPresentationDocument.save(inputFile);
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
