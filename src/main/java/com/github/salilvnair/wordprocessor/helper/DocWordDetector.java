package com.github.salilvnair.wordprocessor.helper;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

abstract class DocWordDetector {

    private String placeholder;
    private String replacementText;


    void findWordsInTable(HWPFDocument doc, String word, String replacementText) {
        this.placeholder = word;
        this.replacementText = replacementText;
   	 	Range range = doc.getRange();
   	 	TableIterator it = new TableIterator(range);  
   	 	while (it.hasNext()) {  
   	 		Table table = (Table) it.next();  
   	 		checkTable(table);
   	 	}
    }

    void findWordsInText(HWPFDocument doc, String word, String replacementText) {
    	this.placeholder = word;
    	Range range = doc.getRange();
    	for (int i = 0; i < range.numSections(); ++i) {
    		Section section = range.getSection(i);
    		checkSection(section);
    	}
    }

    private void checkSection(Section section) {
    	for (int x = 0; x < section.numParagraphs(); x++) {
			Paragraph para = section.getParagraph(x);
			checkInParagraph(para);
    	}
	}

	private void checkTable(Table table) {
    	for (int i = 0; i < table.numRows(); i++) { 
    		TableRow tr = table.getRow(i);
    		checkRow(tr);
    	}
    }

    private void checkRow(TableRow tr ) {
    	for (int j = 0; j < tr.numCells(); j++) {
    		 TableCell td = tr.getCell(j);
    		 checkCell(td);
    	}
    }

    private void checkCell(TableCell td) {
    	 for (int k = 0; k < td.numParagraphs(); k++) {  
    		 Paragraph para = td.getParagraph(k);
    		 checkInParagraph(para);
    	 }
    }

    private void checkInParagraph(Paragraph para) {
    	int totalRun = para.numCharacterRuns();
		 for(int p=0;p<totalRun;p++) {
			 CharacterRun cRun = para.getCharacterRun(p);
			 String runText = cRun.text();
			 if (runText.contains(placeholder)) {
				 cRun.replaceText(placeholder, replacementText);
			 }
		 }
    }
}
