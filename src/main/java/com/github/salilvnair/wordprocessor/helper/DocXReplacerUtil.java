package com.github.salilvnair.wordprocessor.helper;

import com.github.salilvnair.wordprocessor.context.DocumentReplacerContext;
import com.github.salilvnair.wordprocessor.reflect.annotation.PlaceHolder;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.xml.namespace.QName;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.regex.Pattern;

public class DocXReplacerUtil extends DocXWordDetector  implements IDocumentReplacerUtil {
	
    private static final int INITIAL_TEXT_POSITION = 0;

    private String replacementText;
    private String placeHolder;
    private XWPFDocument document;
    private DocumentReplacerContext context;
    
    public void replaceInText(String placeHolder, String replacementText, DocumentReplacerContext context) {
        this.context = context;
        this.replacementText = replacementText;
        this.placeHolder = placeHolder;
        if(context.hasPlaceHolderType()) {
            handleElements(context, placeHolder);
        }
        else {
            findWordsInText(document, placeHolder);
        }
    }

    public void replaceInTable(String placeHolder, String replacementText, DocumentReplacerContext context) {
        this.context = context;
        this.replacementText = replacementText;
        this.placeHolder = placeHolder;
        if(context.hasPlaceHolderType()) {
            context.setReplaceInTable(true);
            handleElements(context, placeHolder);
        }
        else {
            findWordsInTable(document, placeHolder);
        }
    }

    private void handleElements(DocumentReplacerContext context, String placeHolder) {
        PlaceHolder placeHolderType = context.placeHolderType();
        if(placeHolderType.checkbox()) {
            checkCheckboxByTag("true".equals(replacementText));
        }
        else {
            if (context.isReplaceInTable()) {
                findWordsInTable(document, placeHolder);
            }
            else {
                findWordsInText(document, placeHolder);
            }
        }
    }

    public void checkCheckboxByTag(boolean checked) {
        for (XWPFParagraph paragraph : document.getParagraphs()) { //go through all paragraphs
            for (CTSdtRun sdtRun : paragraph.getCTP().getSdtList()) {
                if (W14Checkbox.isW14PlaceHolderCheckbox(sdtRun, this.placeHolder)) {
                    W14Checkbox w14Checkbox = new W14Checkbox(sdtRun);
                    w14Checkbox.checkOrUncheck(checked);
                }
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
            if (sdtPr == null || sdtPr.getTag() == null || placeHolder.equals(sdtPr.getTag().getVal())) {
                return false;
            }
            String declareNameSpaces = "declare namespace w14='http://schemas.microsoft.com/office/word/2010/wordml'";
            XmlObject[] selectedObjects = sdtPr.selectPath(declareNameSpaces + ".//w14:checkbox");
            return selectedObjects.length > 0;
        }
    }
    
    public Object document() {
		return document;
	}

    @Override
    public void onDetection(XWPFRun run) {
    	replaceOnWordDetection(run);
    }

    @Override
    public void onNextDetection(List<XWPFRun> runs, int step) {
    	replaceOnWordDetectionStep(runs, step);
    }
    
    public void init(String docxFile) throws IOException {
    	File file = new File(docxFile);
        init(file);
    }
    
    public void init(File docxFile) throws IOException {
        InputStream inputStream = new FileInputStream(docxFile);
        init(new XWPFDocument(inputStream));
    }

    private void init(XWPFDocument xwpfDoc) {
        if (xwpfDoc == null) throw new NullPointerException();
        document = xwpfDoc;
    }

    private void replaceOnWordDetectionStep(List<XWPFRun> runs, int currentRun) {
        boolean replaced = replaceText(runs.get(currentRun - 1));
        if (replaced) {
            clearText(runs.get(currentRun));
        } else {
        	replaceText(runs.get(currentRun));
        }
        cleanText(runs.get(currentRun + 1));
    }

    private void clearText(XWPFRun run) {
        run.setText("", INITIAL_TEXT_POSITION);
    }

    private void replaceOnWordDetection(XWPFRun run) {
        String replacedText = run.getText(INITIAL_TEXT_POSITION).replaceAll(Pattern.quote(placeHolder), replacementText);
        run.setText(replacedText, INITIAL_TEXT_POSITION);
    }

    private boolean replaceText(XWPFRun run) {
        String text = run.getText(INITIAL_TEXT_POSITION);
        String nextPlaceHolder = placeHolderTail(text, placeHolder);
        if (!nextPlaceHolder.isEmpty()) {
            text = text.replace(nextPlaceHolder, replacementText);
            run.setText(text, INITIAL_TEXT_POSITION);
            return true;
        }
        return false;
    }

    private void cleanText(XWPFRun run) {
        String text = run.getText(INITIAL_TEXT_POSITION);
        String nextPlaceHolder = placeHolderHead(text,placeHolder);
        text = text.replace(nextPlaceHolder, "");
        run.setText(text, INITIAL_TEXT_POSITION);
    }

    private String placeHolderHead(String text, String placeHolder) {
        if (!text.startsWith(placeHolder)) {
            return placeHolderHead(text, placeHolder.substring(1));
        } else {
            return placeHolder;
        }
    }

    private String placeHolderTail(String text, String placeHolder) {
        if (!text.endsWith(placeHolder)) {
            return placeHolderTail(text, placeHolder.substring(0, placeHolder.length() - 1));
        } else {
            return placeHolder;
        }
    }
}
