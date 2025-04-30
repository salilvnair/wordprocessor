package com.github.salilvnair.wordprocessor.helper;

import java.io.IOException;

public interface IDocumentReplacerUtil {
   
	public void replaceInText(String word, String replacementText);
	
	public void replaceInTable(String word, String replacementText);
	
	public void init(String docFile) throws IOException;
	
	public Object document();
}
