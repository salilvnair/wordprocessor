package com.github.salilvnair.wordprocessor.test.main;

import com.github.salilvnair.wordprocessor.writer.WordDocumentWriter;
import com.github.salilvnair.wordprocessor.test.placeholder.bean.TestBean;

public class TestWordProcessor {

	
	public static void main(String[] args) {
		WordDocumentWriter wordDocumentWriter = new WordDocumentWriter();
		TestBean testBean = new TestBean();
		testBean.setTest1(true);
		testBean.setTest2(false);
		testBean.setC("Cat");
		try {
			String template = "/Users/salilvnair/workspace/template.docx";
			wordDocumentWriter
			.placeHolderBeans(testBean)
			.template(template)
			.table()
			.replace()
			.save("/Users/salilvnair/workspace/output_template.docx");
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
	}	
}
