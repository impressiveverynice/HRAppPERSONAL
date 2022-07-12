package kz.jusan.Hr.service.impl;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.util.List;

public class EditorV2 {

    public static void main(String[] args) throws Exception {
        updateDocument("PROFREF1", "44");

    }

    private static void updateDocument(String toreplace, String replacement) throws InvalidFormatException, IOException {

        XWPFDocument doc = new XWPFDocument(OPCPackage.open("C:\\test\\forms.docx")); //CHANGE PATH FOR THE ACTUAL ONE

        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains(toreplace)) {
                                text = text.replace(toreplace, replacement);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains(toreplace)) {
                        text = text.replace(toreplace, replacement);
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains("PIC")) {
                        String imgFile = "C:\\test\\pic.jpg";
                        r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(100), Units.toEMU(100));
                        text = text.replace("PIC", "");
                        r.setText(text, 0);

                    }
                }
            }
        }

        doc.write(new FileOutputStream("C:\\test\\output.docx"));
    }
}
