package com.github.salilvnair.wordprocessor.helper;

import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.List;

interface IDocXWordDetector {

    void onDetection(XWPFRun run);
    void onNextDetection(List<XWPFRun> runs, int step);
}
