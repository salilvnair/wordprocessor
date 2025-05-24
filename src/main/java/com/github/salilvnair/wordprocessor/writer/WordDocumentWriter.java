package com.github.salilvnair.wordprocessor.writer;

import com.github.salilvnair.wordprocessor.context.DocumentReplacerContext;
import com.github.salilvnair.wordprocessor.util.AnnotationUtil;
import com.github.salilvnair.wordprocessor.bean.BaseDocument;
import com.github.salilvnair.wordprocessor.helper.DocReplacerUtil;
import com.github.salilvnair.wordprocessor.helper.DocXReplacerUtil;
import com.github.salilvnair.wordprocessor.helper.IDocumentReplacerUtil;
import com.github.salilvnair.wordprocessor.reflect.annotation.PlaceHolder;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;


public class WordDocumentWriter {
	
	private BaseDocument placeHolderBean;
	private BaseDocument[] placeHolderBeans;
    private String template;
    private final String WORD_FILE_TYPE_DOC = ".doc";
    private final String WORD_FILE_TYPE_DOCX = ".docx";
    private final String PLACEHOLDER_PREFIX = "{{";
    private final String PLACEHOLDER_SUFFIX = "}}";
    private Map<String,Object> placeHolderValueMap;
    private Map<String, PlaceHolder> placeHolderTypeMap;
    private boolean isTableTextReplacement = false;
    private boolean isNormalTextReplacement = false;
	IDocumentReplacerUtil documentReplacerUtil = null;

    public WordDocumentWriter template(String template) {
    	this.template = template; 
    	return this;
    }

    private IDocumentReplacerUtil initDocumentReplacerUtil() throws IOException {
    	if(this.template != null ) {
    		 if (template.endsWith(WORD_FILE_TYPE_DOC)) {
    			 documentReplacerUtil = new DocReplacerUtil();
    			 documentReplacerUtil.init(template);
    		 }
    		 else if (template.endsWith(WORD_FILE_TYPE_DOCX)) {
    			 documentReplacerUtil = new DocXReplacerUtil();
    			 documentReplacerUtil.init(template);
    		 }
    	}
    	return documentReplacerUtil;
    }
    
    private File saveToFile(File file, HWPFDocument document) throws Exception {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file, false);
            document.write(out);
            return file;
        }
		finally {
            if (out != null) {
                out.flush();
                out.close();
            }
        }
    }

    private File saveToFile(File file, XWPFDocument document) throws Exception {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file, false);
            document.write(out);
            return file;
        }
		finally {
            if (out != null) {
                out.flush();
                out.close();
            }
        }
    }
    
    public WordDocumentWriter placeHolderBeans(BaseDocument placeHolderBean) {
    	this.placeHolderBean = placeHolderBean;
    	return this;
    }
    
    public WordDocumentWriter placeHolderBeans(BaseDocument... placeHolderBeans) {
    	this.placeHolderBeans = placeHolderBeans;
    	return this;
    }
        
    private void preparePlaceHolderValueMapFromPlaceHolderBean(BaseDocument baseDocument ) throws InstantiationException, IllegalAccessException, JSONException {
    	PlaceHolder documentPlaceHolder = documentPlaceHolderFromBaseDocument(baseDocument);
    	Set<Field> placeHolderFields = placeHolderFieldsFromBaseDocument(baseDocument);
    	if(documentPlaceHolder!=null && placeHolderFields.isEmpty()) {
    		prepareClassLevelPlaceHolderValueMap(baseDocument);
    	}
    	else if(documentPlaceHolder != null) {
    		prepareClassLevelPlaceHolderValueMap(baseDocument);
    		prepareFieldLevelPlaceHolderValueMap(baseDocument,placeHolderFields);
    	}
    	else {
        	prepareFieldLevelPlaceHolderValueMap(baseDocument,placeHolderFields);
    	}
    }
    
    private void preparePlaceHolderValueMap() throws InstantiationException, IllegalAccessException, JSONException {
    	if(this.placeHolderBean!=null) {
    		preparePlaceHolderValueMapFromPlaceHolderBean(this.placeHolderBean);
    	}
    	else if(this.placeHolderBeans!=null) {
    		for(BaseDocument placeHolderBean:placeHolderBeans) {
    			preparePlaceHolderValueMapFromPlaceHolderBean(placeHolderBean);
    		}
    	}
    }
    
    private void prepareFieldLevelPlaceHolderValueMap(BaseDocument baseDocument,
													  Set<Field> placeHolderFields) throws JSONException {
    	if(placeHolderValueMap==null) {
    		placeHolderValueMap = new HashMap<>();
    	}
		if (placeHolderTypeMap == null) {
			placeHolderTypeMap = new HashMap<>();
		}
		for(Field placeHolderField:placeHolderFields){
			PlaceHolder documentPlaceHolder = placeHolderField.getAnnotation(PlaceHolder.class);
			if(documentPlaceHolder == null) {
				continue;
			}
			String placeHolderKey = documentPlaceHolder.value();
			placeHolderTypeMap.put(placeHolderKey, documentPlaceHolder);
			Gson gson = new GsonBuilder().setDateFormat("E MMM dd hh:mm:ss Z yyyy").create();
			String jsonString = gson.toJson(baseDocument);
			JSONObject jasonObject = new JSONObject(jsonString);
			String key = placeHolderField.getName();
			Object hasValue=jasonObject.opt(key);
			if(hasValue!=null){
				Object jasonValue = jasonObject.get(key);
				placeHolderValueMap.put(placeHolderKey, jasonValue);
			}
			if(documentPlaceHolder.nonNull()) {
				if(hasValue==null){
					placeHolderValueMap.put(placeHolderKey, documentPlaceHolder.replaceNullWith());
				}
				else {
					Object jasonValue = jasonObject.get(key);
					if(jasonValue==null) {
						placeHolderValueMap.put(placeHolderKey, documentPlaceHolder.replaceNullWith());
					}
				}
			}		
		}
	}

	private void prepareClassLevelPlaceHolderValueMap(BaseDocument baseDocument) throws JSONException {
    	if(placeHolderValueMap==null) {
    		placeHolderValueMap = new HashMap<>();
    	}
		Field[] fields = baseDocument.getClass().getDeclaredFields();
		Gson gson = new GsonBuilder().setDateFormat("E MMM dd hh:mm:ss Z yyyy").create();
		String jsonString = gson.toJson(baseDocument);
		JSONObject jasonObject = new JSONObject(jsonString);
    	for (Field field : fields) {
    		if(field.isAnnotationPresent(PlaceHolder.class)) {
    			continue;
    		}
    		String key = field.getName();
			Object hasValue=jasonObject.opt(key);
			if(hasValue!=null) {
				Object jasonValue = jasonObject.get(key);
				String placeHolderKey = PLACEHOLDER_PREFIX+key+PLACEHOLDER_SUFFIX;
				placeHolderValueMap.put(placeHolderKey, jasonValue);
			}
    	}
	}

	private  Set<Field> placeHolderFieldsFromBaseDocument(BaseDocument baseDocument) {
    	return AnnotationUtil.findAnnotatedFields(baseDocument.getClass(), PlaceHolder.class);
    }
    
	private  PlaceHolder documentPlaceHolderFromBaseDocument(BaseDocument baseDocument) {
		PlaceHolder documentPlaceHolder = null;
		if(baseDocument.getClass().isAnnotationPresent(PlaceHolder.class)){
			documentPlaceHolder = baseDocument.getClass().getAnnotation(PlaceHolder.class);
		}
		return documentPlaceHolder;
	}
	
	public WordDocumentWriter table() throws InstantiationException, IllegalAccessException, JSONException, IOException {
		initProcessor();
		this.isTableTextReplacement = true;
		return this;
	}
	
	private void initProcessor() throws InstantiationException, IllegalAccessException, JSONException, IOException {
		initDocumentReplacerUtil();
		preparePlaceHolderValueMap();
	}

	public WordDocumentWriter text() throws InstantiationException, IllegalAccessException, JSONException, IOException {
		initProcessor();
		this.isNormalTextReplacement = true;
		return this;
	}
	
	public WordDocumentWriter replace() {
		if(this.isNormalTextReplacement) {
			for(String placeHolder : placeHolderValueMap.keySet()) {
				Object replacementText = placeHolderValueMap.get(placeHolder);
				PlaceHolder placeHolderType = placeHolderTypeMap.get(placeHolder);
				DocumentReplacerContext context = DocumentReplacerContext
													.builder()
													.placeHolderText(placeHolder)
													.placeHolderValueMap(placeHolderValueMap)
													.placeHolderType(placeHolderType)
													.placeHolderTypeMap(placeHolderTypeMap)
													.build();
				documentReplacerUtil.replaceInText(placeHolder, replacementText+"", context);
			}
		}
		else if (this.isTableTextReplacement) {
			for(String placeHolder : placeHolderValueMap.keySet()) {
				Object replacementText = placeHolderValueMap.get(placeHolder);
				PlaceHolder placeHolderType = placeHolderTypeMap.get(placeHolder);
				DocumentReplacerContext context = DocumentReplacerContext
													.builder()
													.placeHolderText(placeHolder)
													.placeHolderValueMap(placeHolderValueMap)
													.placeHolderTypeMap(placeHolderTypeMap)
													.placeHolderType(placeHolderType)
													.build();
				documentReplacerUtil.replaceInTable(placeHolder, replacementText+"", context);
			}
		}
		return this;
	}
	
	public void save(String fileNameWithPath) throws Exception {
		File file = new File(fileNameWithPath);
		save(file);
	}
	
	public void save(File file) throws Exception {
		Object document = documentReplacerUtil.document();
		if(document instanceof XWPFDocument xwpfDocument) {
            saveToFile(file, xwpfDocument);
		}
		else if(document instanceof HWPFDocument hwpfDocument) {
            saveToFile(file, hwpfDocument);
		}
	}
	
	public Object generateWordDocument() {
		return documentReplacerUtil.document();
	}
	
	public XWPFDocument generateDocX() {
		Object document = documentReplacerUtil.document();
		if(document instanceof XWPFDocument) {
            return (XWPFDocument) document;
		}
		return null;
	}
	
	public HWPFDocument generateDoc() {
		Object document = documentReplacerUtil.document();
		if(document instanceof HWPFDocument) {
			return (HWPFDocument) document;
		}
		return null;
	}
}
