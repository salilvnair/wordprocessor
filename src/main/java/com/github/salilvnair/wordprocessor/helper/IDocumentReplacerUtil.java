package com.github.salilvnair.wordprocessor.helper;

import com.github.salilvnair.wordprocessor.context.DocumentReplacerContext;

import java.io.IOException;

public interface IDocumentReplacerUtil {
   
	public void replaceInText(String word, String replacementText, DocumentReplacerContext context);
	
	public void replaceInTable(String word, String replacementText, DocumentReplacerContext context);
	
	public void init(String docFile) throws IOException;
	
	public Object document();
}
