package com.github.salilvnair.wordprocessor.helper;

import com.github.salilvnair.wordprocessor.constant.WordProcessorConstant;
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
    private final DocXCheckboxDetector docXCheckboxDetector;

    public DocXReplacerUtil() {
        docXCheckboxDetector = new DocXCheckboxDetector();
    }
    
    public void replaceInText(String placeHolder, String replacementText, DocumentReplacerContext context) {
        this.context = context;
        this.replacementText = replacementText;
        this.placeHolder = placeHolder;
        if(context.hasPlaceHolderType() && context.placeHolderType().checkbox()) {
            docXCheckboxDetector.findCheckboxInParagraph(document, placeHolder, WordProcessorConstant.TRUE.equals(replacementText));
        }
        else {
            findWordsInText(document, placeHolder);
        }
    }

    public void replaceInTable(String placeHolder, String replacementText, DocumentReplacerContext context) {
        this.context = context;
        this.replacementText = replacementText;
        this.placeHolder = placeHolder;
        if(context.hasPlaceHolderType() && context.placeHolderType().checkbox()) {
            docXCheckboxDetector.findCheckboxInTable(document, placeHolder, WordProcessorConstant.TRUE.equals(replacementText));
        }
        else {
            findWordsInTable(document, placeHolder);
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
