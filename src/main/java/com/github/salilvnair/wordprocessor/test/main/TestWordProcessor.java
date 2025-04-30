package com.github.salilvnair.wordprocessor.test.main;

import com.github.salilvnair.wordprocessor.writer.WordDocumentWriter;
import com.github.salilvnair.wordprocessor.test.placeholder.bean.TestCompetitiveVerificationForm;
import com.github.salilvnair.wordprocessor.test.placeholder.bean.TestOptyBean;

public class TestWordProcessor {

	
	public static void main(String[] args) {
		WordDocumentWriter wordDocumentWriter = new WordDocumentWriter();
		TestCompetitiveVerificationForm cmp = new TestCompetitiveVerificationForm();
		cmp.setCustomerName("Sherlock Holmes");
		cmp.setTitle("Mr");
		cmp.setAddress("221B Baker Street");
		cmp.setCity("Marylebone");
		cmp.setState("London");
		cmp.setCustomerContract("XXX12345");
		cmp.setPhone("12345");
		TestOptyBean optyBean = new TestOptyBean();
		optyBean.setOptyNew(true);
		try {
			String template = "C:\\MyHDD\\Z\\playground\\wordprocessor\\input_template.docx";
//			Object document = wordProcessorBuilder
//			.placeHolderBeans(cmp)
//			.setTemplate(template)
//			.table()
//			.replace()
//			.generateWordDocument();
			wordDocumentWriter
			.placeHolderBeans(cmp,optyBean)
			.template(template)
			.table()
			.replace()
			.save("C:\\MyHDD\\Z\\playground\\wordprocessor\\output_21_11.docx");
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
	}	
}
