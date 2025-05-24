package com.github.salilvnair.wordprocessor.context;

import com.github.salilvnair.wordprocessor.reflect.annotation.PlaceHolder;
import lombok.*;

import java.util.Map;

@Getter
@Setter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DocumentReplacerContext {
    private String placeHolderText;
    private PlaceHolder placeHolderType;
    private Map<String,Object> placeHolderValueMap;
    private Map<String, PlaceHolder> placeHolderTypeMap;
    private boolean replaceInTable;

    public boolean hasPlaceHolderType() {
        return placeHolderType != null;
    }
    public PlaceHolder placeHolderType() {
        return placeHolderType;
    }
}
