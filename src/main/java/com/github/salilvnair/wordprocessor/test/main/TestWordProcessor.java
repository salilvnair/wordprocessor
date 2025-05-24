package com.github.salilvnair.wordprocessor.test.main;

import com.github.salilvnair.wordprocessor.writer.WordDocumentWriter;
import com.github.salilvnair.wordprocessor.test.placeholder.bean.TestBean;

public class TestWordProcessor {

	
	public static void main(String[] args) {
		WordDocumentWriter wordDocumentWriter = new WordDocumentWriter();
		TestBean testBean = new TestBean();
		testBean.setTest1(false);
		testBean.setTest2(false);
		testBean.setA("Acha");
		testBean.setB("Bacha");
		testBean.setC("Chacha");
		try {
			String template = "C:\\Users\\sn2527\\workspace\\template.docx";
			wordDocumentWriter
			.placeHolderBeans(testBean)
			.template(template)
			.table()
			.replace()
			.save("C:\\Users\\sn2527\\workspace\\output_21_12.docx");
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
	}	
}
