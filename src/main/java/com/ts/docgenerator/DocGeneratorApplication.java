package com.ts.docgenerator;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.swing.text.TableView;
import java.io.*;
import java.math.BigInteger;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class DocGeneratorApplication {

    public static void main(String[] args) throws IOException, URISyntaxException {
        SpringApplication.run(DocGeneratorApplication.class, args);
        String resourcePath = "DemoPOIWord1.docx";
        Path templatePath = Paths.get(DocGeneratorApplication.class.getClassLoader().getResource(resourcePath).toURI());
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));
        doc = replaceTable(doc);
        //doc = replaceTextFor(doc, "${placeholder1}", "Techmonks");
        saveWord("D:\\document.docx", doc);
//        try {
//            prepareDoc();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }

    static void prepareDoc() throws IOException {
        XWPFDocument document = new XWPFDocument();
        /* Create Paragraph */
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("At ${placeholder1}, we strive hard to " +
                "provide quality tutorials for self-learning " +
                "purpose in the domains of Academics, Information " +
                "Technology, Management and Computer Programming Languages.");
        run.addCarriageReturn();
        XWPFParagraph tableParagraph = document.createParagraph();
        XWPFRun tableRun = tableParagraph.createRun();
        tableRun.setText("${tblReleaseItems}");

        /*
         * Create table with formating text in XWPFTable
         */
        XWPFTable table = document.createTable();
        XWPFTableRow tableRow1 = table.getRow(0);


        /* Removing paragraph before setText for cell*/
        tableRow1.getCell(0).removeParagraph(0);

        /* Create cellRun to format text*/
        XWPFRun cellRun1 = tableRow1.getCell(0).addParagraph().createRun();
        cellRun1.setText("Sample Text");
        cellRun1.addBreak();
        cellRun1.setText("Sample Text");
        cellRun1.addBreak();
        cellRun1.setText("Sample Text");
        cellRun1.addBreak();
        cellRun1.setFontSize(12);
        cellRun1.setBold(true);

        /* Insert image */
//		Path imagePath = Paths.get(ClassLoader.getSystemResource("bachkhoa.png").toURI());
//		cellRun1.addPicture(Files.newInputStream(imagePath),
//				XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(),
//				Units.toEMU(40), Units.toEMU(60));

        /* Set Alignment "CENTER" of data in cell */
        XWPFTableCell cell1 = tableRow1.getCell(0);
        XWPFParagraph paragraph1 = cell1.getParagraphs().get(0);
        paragraph1.setAlignment(ParagraphAlignment.CENTER);

        /* Cell 2 */
        XWPFRun cellRun2 = tableRow1.addNewTableCell().addParagraph().createRun();
        tableRow1.getCell(1).removeParagraph(0);
        cellRun2.setText("Sample Text");
        cellRun2.setFontSize(12);
        cellRun2.setBold(true);

        /* Set width for table*/
        table.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(10000));

        /*
         * Write the Document in file system  Y
         * You can reset path of file.
         * Here, I use "C:/Users/User/Music/Desktop/DemoPOIWord.docx
         */
        FileOutputStream out = new FileOutputStream(new File("D:/DemoPOIWord1.docx"));
        //OutputStream outputStream = new
        document.write(out);
        out.close();
        System.out.println("DemoPOIWord.docx written successfully");
    }

    private static XWPFDocument replaceTextFor(XWPFDocument doc, String findText, String replaceText) {
        doc.getParagraphs().forEach(p -> {
            p.getRuns().forEach(run -> {
                String text = run.text();
                if (text.contains(findText)) {
                    run.setText(text.replace(findText, replaceText), 0);
                }
            });
        });

        return doc;
    }

    private static XWPFDocument replaceTable(XWPFDocument doc) {
        XWPFTable table = null;
        long count = 0;
        System.out.println(doc.getParagraphs().size());
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            String find = "${tblReleaseItems}";
            TextSegment found =
                    paragraph.searchText(find, new PositionInParagraph());
            if (found != null) {
                count++;
                if (found.getBeginRun() == found.getEndRun()) {
                    // whole search string is in one Run
                    XmlCursor cursor = paragraph.getCTP().newCursor();
                    table = doc.insertNewTbl(cursor);

                    XWPFRun run = runs.get(found.getBeginRun());
                    // Clear the "%TABLE" from doc
                    String runText = run.getText(run.getTextPosition());
                    String replaced = runText.replace(find, "");
                    run.setText(replaced, 0);
                } else {
                    // The search string spans over more than one Run
                    StringBuilder b = new StringBuilder();
                    for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
                        XWPFRun run = runs.get(runPos);
                        b.append(run.getText(run.getTextPosition()));
                    }
                    String connectedRuns = b.toString();
                    XmlCursor cursor = paragraph.getCTP().newCursor();
                    table = doc.insertNewTbl(cursor);
                    //fillInTable(table);
                    String replaced = connectedRuns.replace(find, ""); // Clear search text

                    // The first Run receives the replaced String of all connected Runs
                    XWPFRun partOne = runs.get(found.getBeginRun());
                    partOne.setText(replaced, 0);
                    // Removing the text in the other Runs.
                    for (int runPos = found.getBeginRun() + 1; runPos <= found.getEndRun(); runPos++) {
                        XWPFRun partNext = runs.get(runPos);
                        partNext.setText("", 0);
                    }
                }
            }
        }
        fillInTable(table);
        //return count;
        return doc;
        //https://stackoverflow.com/questions/22268898/replacing-a-text-in-apache-poi-xwpf
    }

    private static void fillInTable(XWPFTable tableInfo) {
//        XWPFTableRow headerRow = table.getRow(0);
//        addTableRowCell(headerRow, 0, "Req.Id", true);
//        addTableRowCell(headerRow, 1, "Title", true);
//        addTableRowCell(headerRow, 2, "Status", true);
//        addTableRowCell(headerRow, 3, "Assigned To", true);
//        addTableRowCell(headerRow, 4, "Reported By", true);

//        XWPFTableRow childRow1 = table.insertNewTableRow(1);
//        addTableRowCell(childRow1, 0, "RAZE-1", false);
//        addTableRowCell(childRow1, 1, "UI Dashboard", false);
//        addTableRowCell(childRow1, 2, "Completed", false);
//        addTableRowCell(childRow1, 3, "Krishna Varma", false);
//        addTableRowCell(childRow1, 4, "Krishna Varma", false);


        /* Insert image */
//		Path imagePath = Paths.get(ClassLoader.getSystemResource("bachkhoa.png").toURI());
//		cellRun1.addPicture(Files.newInputStream(imagePath),
//				XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(),
//				Units.toEMU(40), Units.toEMU(60));

        /* Set Alignment "CENTER" of data in cell */
//        XWPFTableCell cell1 = tableRow1.getCell(0);
//        XWPFParagraph paragraph1 = cell1.getParagraphs().get(0);
//        paragraph1.setAlignment(ParagraphAlignment.CENTER);

        XWPFTable table = tableInfo;
        XWPFTableRow tableRowOne = table.getRow(0);
        //tableRowOne.getCell(0).setText("col one, row one");
        addTableRowCell(tableRowOne, 0,"Req. Id", true);
        addTableRowCell(tableRowOne, 1,"Title", true);
        addTableRowCell(tableRowOne, 2,"Status", true);
        addTableRowCell(tableRowOne, 3, "Assigned To", true);
        addTableRowCell(tableRowOne, 4, "Reported By", true);

        XWPFTableRow childRow1 = table.createRow();
        addTableRowCell(childRow1, 0, "RAZE-1", false);
        addTableRowCell(childRow1, 1, "UI Dashboard", false);
        addTableRowCell(childRow1, 2, "Completed", false);
        addTableRowCell(childRow1, 3, "Krishna Varma", false);
        addTableRowCell(childRow1, 4, "Krishna Varma", false);

        XWPFTableRow childRow2 = table.createRow();
        addTableRowCell(childRow2, 0, "RAZE-5", false);
        addTableRowCell(childRow2, 1, "External Integration", false);
        addTableRowCell(childRow2, 2, "Completed", false);
        addTableRowCell(childRow2, 3, "Krishna Varma", false);
        addTableRowCell(childRow2, 4, "Krishna Varma", false);

        XWPFTableRow childRow3 = table.createRow();
        addTableRowCell(childRow3, 0, "RAZE-1", false);
        addTableRowCell(childRow3, 1, "Include Authentication in API", false);
        addTableRowCell(childRow3, 2, "Completed", false);
        addTableRowCell(childRow3, 3, "Krishna Varma", false);
        addTableRowCell(childRow3, 4, "Krishna Varma", false);

        XWPFTableRow childRow4 = table.createRow();
        addTableRowCell(childRow4, 0, "RAZE-6", false);
        addTableRowCell(childRow4, 1, "Add additional items to release note", false);
        addTableRowCell(childRow4, 2, "Completed", false);
        addTableRowCell(childRow4, 3, "Krishna Varma", false);
        addTableRowCell(childRow4, 4, "Krishna Varma", false);

        //create second row
//        XWPFTableRow tableRowTwo = table.createRow();
//        addTableRowCell(tableRowTwo, 0, "col one, row two", false);
//        addTableRowCell(tableRowTwo, 1, "col two, row two", false);
//        addTableRowCell(tableRowTwo, 2, "col three, row two", false);

        //create third row
//        XWPFTableRow tableRowThree = table.createRow();
//        addTableRowCell(tableRowThree, 0, "col one, row three", false);
//        addTableRowCell(tableRowThree, 1, "col two, row three", false);
//        addTableRowCell(tableRowThree, 2, "col three, row three", false);


//        tableRowThree.getCell(0).setText("col one, row three");
//        tableRowThree.getCell(1).setText("col two, row three");
//        tableRowThree.getCell(2).setText("col three, row three");/* Set width for table*/
        table.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(10000));
    }

    private static void addTableRowCell(XWPFTableRow xwpfTableRow, Integer position, String valueText, Boolean headerRow) {
        if (headerRow) {
            XWPFRun cellRun2 = null;
            if(position == 0){
                cellRun2 = xwpfTableRow.getCell(position).addParagraph().createRun();
            }
            else {
                cellRun2 = xwpfTableRow.addNewTableCell().addParagraph().createRun();
            }
            xwpfTableRow.getCell(position).removeParagraph(0);
            cellRun2.setText(valueText);
            cellRun2.setFontSize(12);
            cellRun2.setBold(true);
        } else {
            xwpfTableRow.getCell(position).setText(valueText);
        }
    }

    private static void saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            out.close();
        }
    }
}
