package com.github.salilvnair.wordprocessor.test.placeholder.bean;

import com.github.salilvnair.wordprocessor.bean.BaseDocument;
import com.github.salilvnair.wordprocessor.reflect.annotation.PlaceHolder;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TestOptyBean extends BaseDocument {
	//if the PlaceHolder is annotated at the field level then the placeholder text for
	//the document will be expected as whatever given in the value attribute
	//checkbox=true if set explicitly then it is mandatory to give its type as boolean
	//for all the fields whose value is true in that case will be replaced by 
	//checkboxCheckedValue's default value or by whatever given as the attribute value
	//and false values will be replaced by checkboxUnheckedValue's default value or by whatever given
	//as the attribute value
	@PlaceHolder(value="{{opty.newYn}}",checkbox=true)
	private boolean optyNew;
	
	@PlaceHolder(value="{{opty.winback}}",checkbox=true)
	private boolean winback;
	
	
	@PlaceHolder(value="{{opty.mig}}",checkbox=true)
	private boolean mig;
	
	
	@PlaceHolder(value="{{opty.ret}}",checkbox=true)
	private boolean ret;
	
	
}
