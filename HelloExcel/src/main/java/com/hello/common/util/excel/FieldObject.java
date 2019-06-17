package com.hello.common.util.excel;

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * 描述：excel 字段属性
 * Created by jay on 2017-9-18.
 */
@Data
public class FieldObject {

    private String name;

    private boolean lockBoolean;

    private ExcelDataEnums format;

    private int index;

    private List<String> constantSelectList;

    private Map<String, List<String>> map;

    private String parentName;

    private String pointOut;

    private Class<? extends ExcelSelectInterface> returnSelectDataClass;

    private Class<? extends ExcelSelectMapInterface> returnSelectMapClass;

    public boolean getLockBoolean() {
        return lockBoolean;
    }
}