package com.ts.docgenerator.controller;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.springframework.http.*;
import org.springframework.util.StreamUtils;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigInteger;
import java.util.List;

@CrossOrigin
@RestController
@RequestMapping("v1/docs")
public class DocController {

    private static final String APPLICATION_MS_WORD_VALUE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    @GetMapping("download")
    public HttpEntity download(HttpServletResponse res) throws IOException {
        XWPFDocument document = constructDocument();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        document.write(baos);
        byte[] byteArray = baos.toByteArray();
        return ResponseEntity.ok()
                .contentLength(byteArray.length)
                .header(HttpHeaders.CONTENT_TYPE, APPLICATION_MS_WORD_VALUE)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + "File.docx")
                             .body(byteArray);
    }

    private XWPFDocument constructDocument(){
        XWPFDocument document = new XWPFDocument();
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "This is header";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

        /* Create Paragraph */
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("At ${placeholder1}, we strive hard to " +
                "provide quality tutorials for self-learning " +
                "purpose in the domains of Academics, Information " +
                "Technology, Management and Computer Programming Languages.");

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
        return document;
    }
}
