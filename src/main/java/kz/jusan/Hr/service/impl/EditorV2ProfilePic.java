package kz.jusan.Hr.service.impl;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.Units;

import org.apache.poi.xwpf.usermodel.*;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;

import java.math.BigInteger;

public class EditorV2ProfilePic {

    public static void main(String[] args) throws Exception {

        XWPFDocument doc = new XWPFDocument(OPCPackage.open("C:\\test\\forms.docx"));

        // the body content
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("The Body:");

        paragraph = doc.createParagraph();
        run = paragraph.createRun();
        run.setText("Lorem ipsum....");

        // create header start
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);

        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);

        CTTabStop tabStop = paragraph.getCTP().getPPr().addNewTabs().addNewTab();
        tabStop.setVal(STTabJc.RIGHT);
        int twipsPerInch = 1440;
        tabStop.setPos(BigInteger.valueOf(6 * twipsPerInch));

        run = paragraph.createRun();
        run.setText("The Header:");
        run.addTab();

        run = paragraph.createRun();
        String imgFile = "C:\\test\\pic.jpg";
        run.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(100), Units.toEMU(100));


        // create footer start
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);

        paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        run = paragraph.createRun();
        run.setText("The Footer:");


        doc.write(new FileOutputStream("C:\\test\\output.docx"));

    }
}