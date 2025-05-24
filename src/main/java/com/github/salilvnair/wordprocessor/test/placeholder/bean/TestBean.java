package com.github.salilvnair.wordprocessor.test.placeholder.bean;

import com.github.salilvnair.wordprocessor.bean.BaseDocument;
import com.github.salilvnair.wordprocessor.reflect.annotation.PlaceHolder;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TestBean extends BaseDocument {
	//if the PlaceHolder is annotated at the field level then the placeholder text for
	//the document will be expected as whatever given in the value attribute
	//checkbox=true if set explicitly then it is mandatory to give its type as boolean
	@PlaceHolder(value="test1",checkbox=true)
	private boolean test1;

	@PlaceHolder(value="test2",checkbox=true)
	private boolean test2;

	@PlaceHolder(value="{{a}}",nonNull=true,replaceNullWith=" ")
	private String a;

	@PlaceHolder(value="{{b}}",nonNull=true,replaceNullWith=" ")
	private String b;

	@PlaceHolder(value="{{c}}",nonNull=true,replaceNullWith=" ")
	private String c;
	
}
