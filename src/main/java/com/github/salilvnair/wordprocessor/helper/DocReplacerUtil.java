package com.github.salilvnair.wordprocessor.helper;

import com.github.salilvnair.wordprocessor.context.DocumentReplacerContext;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class DocReplacerUtil extends DocWordDetector implements IDocumentReplacerUtil{
	
	private HWPFDocument document;
	private DocumentReplacerContext context;
	
	public void init(HWPFDocument document) {
		this.document = document;
	}
	
	public Object document() {
		return document;
	}
	
	public void init(String docFile) throws IOException {
		File file = new File(docFile);
		init(file);
	}
	
	public void init(File docFile) throws IOException {
		InputStream inputStream = new FileInputStream(docFile);
		this.document = new HWPFDocument(inputStream);
	}
	
    public void replaceInText(String word, String replacementText, DocumentReplacerContext context) {
		this.context = context;
    	findWordsInText(document, word, replacementText);
	}
	
	public void replaceInTable(String word, String replacementText, DocumentReplacerContext context) {
		this.context = context;
		findWordsInTable(document, word, replacementText);
	}
}
