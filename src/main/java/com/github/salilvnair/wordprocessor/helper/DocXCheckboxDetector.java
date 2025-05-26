package com.github.salilvnair.wordprocessor.helper;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import javax.xml.namespace.QName;
import java.util.List;

public class DocXCheckboxDetector {
    private String tag;
    private boolean checked;


    void findCheckboxInTable(XWPFDocument doc, String tag, boolean checked) {
        this.tag = tag;
        this.checked = checked;
        for (XWPFTable t : doc.getTables()) {
            checkTable(t);
        }
    }

    void findCheckboxInParagraph(XWPFDocument doc, String tag, boolean checked) {
        this.tag = tag;
        this.checked = checked;
        for (XWPFParagraph p : doc.getParagraphs()) {
            if (p != null && !p.getRuns().isEmpty()) {
                checkInParagraph(p, checked);
            }
        }
    }

    private void checkTable(XWPFTable t) {
        if (t.getRows() == null) {
        	return;
        }
        for (XWPFTableRow r : t.getRows()) {
            checkRow(r);
        }
    }

    private void checkRow(XWPFTableRow r) {
        if (r.getTableCells() == null) {
        	return;
        }
        for (XWPFTableCell cell : r.getTableCells()) {
            checkCell(cell);
        }
    }

    private void checkCell(XWPFTableCell cell) {
        List<IBodyElement> cellBodyElements = cell.getBodyElements();
        if (cellBodyElements == null) {
        	return;
        }
        for(IBodyElement cellBodyElement:cellBodyElements) {
        	if(cellBodyElement instanceof XWPFTable t) {
                checkTable(t);
        	}
        	else if(cellBodyElement instanceof XWPFParagraph p) {
                if (!p.getRuns().isEmpty()) {
                    checkInParagraph(p, checked);
                }
        	}
        }
    }

    private void checkInParagraph(XWPFParagraph paragraph, boolean checked) {
        for (CTSdtRun sdtRun : paragraph.getCTP().getSdtList()) {
            if (W14Checkbox.isW14PlaceHolderCheckbox(sdtRun, this.tag)) {
                W14Checkbox w14Checkbox = new W14Checkbox(sdtRun);
                w14Checkbox.checkOrUncheck(checked);
            }
        }
    }

    static class W14Checkbox {
        CTSdtRun sdtRun;
        CTSdtContentRun sdtContentRun = null;
        XmlObject w14CheckboxChecked = null;

        W14Checkbox(CTSdtRun sdtRun) {
            this.sdtRun = sdtRun;
            this.sdtContentRun = sdtRun.getSdtContent();
            String declareNameSpaces = "declare namespace w14='http://schemas.microsoft.com/office/word/2010/wordml'";
            XmlObject[] selectedObjects = sdtRun.getSdtPr().selectPath(declareNameSpaces + ".//w14:checkbox/w14:checked");
            if (selectedObjects.length > 0) {
                this.w14CheckboxChecked = selectedObjects[0];
            }
        }

        CTSdtContentRun content() {
            return this.sdtContentRun;
        }

        XmlObject w14CheckboxChecked() {
            return this.w14CheckboxChecked;
        }

        boolean checked() {
            XmlCursor cursor = this.w14CheckboxChecked.newCursor();
            String val = cursor.getAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"));
            return "1".equals(val) || "true".equals(val);
        }

        void checkOrUncheck(boolean checked) {
            if (checked) {
                check();
            }
            else {
                uncheck();
            }
        }

        void check() {
            XmlCursor cursor = this.w14CheckboxChecked.newCursor();
            String val = "1";
            cursor.setAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"), val);
            cursor.close();
            CTText t = this.sdtContentRun.getRArray(0).getTArray(0);
            String content = "☒";
            t.setStringValue(content);
        }
        void uncheck() {
            XmlCursor cursor = this.w14CheckboxChecked.newCursor();
            String val = "0";
            cursor.setAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"), val);
            cursor.close();
            CTText t = this.sdtContentRun.getRArray(0).getTArray(0);
            String content = "☐";
            t.setStringValue(content);
        }

        static boolean isW14PlaceHolderCheckbox(CTSdtRun sdtRun, String placeHolder) {
            if (sdtRun == null) {
                return false;
            }
            CTSdtPr sdtPr = sdtRun.getSdtPr();
            if (sdtPr == null || sdtPr.getTag() == null || !placeHolder.equals(sdtPr.getTag().getVal())) {
                return false;
            }
            String declareNameSpaces = "declare namespace w14='http://schemas.microsoft.com/office/word/2010/wordml'";
            XmlObject[] selectedObjects = sdtPr.selectPath(declareNameSpaces + ".//w14:checkbox");
            return selectedObjects.length > 0 && placeHolder.equals(sdtPr.getTag().getVal());
        }
    }
}
