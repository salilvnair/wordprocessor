package com.github.salilvnair.wordprocessor.helper;

import org.apache.poi.xwpf.usermodel.*;

import java.util.List;

abstract class DocXWordDetector implements IDocXWordDetector {

    private static final int DEFAULT_POSITION = 0;

    private String placeholder;


    void findWordsInTable(XWPFDocument doc, String word) {
        this.placeholder = word;
        for (XWPFTable t : doc.getTables()) {
            checkTable(t);
        }
    }

    void findWordsInText(XWPFDocument doc, String word) {
        this.placeholder = word;
        for (XWPFParagraph p : doc.getParagraphs()) {
            if (p != null && !p.getRuns().isEmpty()) {
                checkInParagraph(p);
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
                    checkInParagraph(p);
                }
        	}
        }
    }

    private void checkInParagraph(XWPFParagraph p) {
        List<XWPFRun> runs = p.getRuns();
        int prevRunIndex = -1;
        for (int runIndex = 0; runIndex < runs.size(); runIndex++) {
            XWPFRun run = p.getRuns().get(runIndex);
            if (run != null && run.getText(DEFAULT_POSITION) != null) {
                String text = run.getText(DEFAULT_POSITION);               
                if (text.contains(placeholder)) {
                	onDetection(run);
                	prevRunIndex = runIndex;
                } else if (hasText(runs, runIndex)
                        && !nextText(runs, runIndex).contains(placeholder)
                        && wordDetected(runs, prevRunIndex, runIndex)) {
                	onNextDetection(runs, runIndex);
                }
            }
        }
    }

    private boolean wordDetected(List<XWPFRun> runs, int prevRunIndex, int runIndex) {
        return notFirstRun(runIndex)
                && prevRunHasText(runs, runIndex)
                && validatePrevRun(prevRunIndex, runIndex)
                && lastRunText(runs, runIndex).contains(placeholder);
    }

    private boolean validatePrevRun(int prevRunIndex, int runIndex) {
        return prevRunIndex != runIndex - 1;
    }

    private String lastRunText(List<XWPFRun> runs, int runIndex) {
        String text = runs.get(runIndex).getText(DEFAULT_POSITION);
        return lastText(runs, runIndex, text) + nextText(runs, runIndex);
    }

    private boolean hasText(List<XWPFRun> runs, int runIndex) {
        return runs.size() > runIndex + 1
                && runs.get(runIndex + 1).getText(DEFAULT_POSITION) != null
                && !runs.get(runIndex + 1).getText(DEFAULT_POSITION).isEmpty();
    }

    private String nextText(List<XWPFRun> runs, int runIndex) {
        return runs.get(runIndex + 1).getText(DEFAULT_POSITION);
    }

    private String lastText(List<XWPFRun> runs, int runIndex, String text) {
        return runs.get(runIndex - 1).getText(DEFAULT_POSITION) + text;
    }

    private boolean prevRunHasText(List<XWPFRun> runs, int runIndex) {
        return runs.get(runIndex - 1).getText(DEFAULT_POSITION) != null
                && !runs.get(runIndex - 1).getText(DEFAULT_POSITION).isEmpty();
    }

    private boolean notFirstRun(int runIndex) {
        return runIndex > 0;
    }

}
