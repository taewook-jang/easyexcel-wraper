package com.wuyue.excel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.util.StringUtils;
import com.wuyue.excel.ExcelRow.BindingError;
import com.wuyue.excel.validate.HibernateValidator;
import com.wuyue.excel.validate.NotDuplicateValidator;
import lombok.Getter;
import lombok.Setter;

import javax.validation.ConstraintViolation;
import javax.validation.groups.Default;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class BaseEventListener<T extends ExcelRow> extends AnalysisEventListener<T> {

    @Getter
    private List<T> rows;

    private Field[] fields;

    private Class<T> tClass;

    @Setter
    private int headLineMun;

    private NotDuplicateValidator notDuplicateValidator;

    public BaseEventListener(Class<T> tClass) {
        fields = tClass.getDeclaredFields();
        rows = new ArrayList<>();
        this.tClass = tClass;
        notDuplicateValidator = new NotDuplicateValidator();
    }

    @Override
    public void invoke(T r, AnalysisContext analysisContext) {
        recoverEmptyRow(analysisContext.getCurrentRowNum() - 1, rows.size());
        validate(r);

        Map<Integer, Head> headMap = analysisContext.currentReadHolder().excelReadHeadProperty().getHeadMap();

        Map<String, Integer> columnMetaMap = new HashMap<>();

        for (Map.Entry<Integer, Head> entry : headMap.entrySet()) {
            Head head = entry.getValue();
            columnMetaMap.put(head.getFieldName(), head.getColumnIndex());
        }


        /*if (r.getValidateCode() == 2) {
            r.setRowNum(analysisContext.readRowHolder().getRowIndex());
            //r.setColNum(analysisContext.get);
            rows.add(r);
        }*/

        r.setRowNum(analysisContext.readRowHolder().getRowIndex());
        //r.setColNum(analysisContext.get);
        rows.add(r);

    }



    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        recoverEmptyRow(analysisContext.getCurrentRowNum() - 1, rows.size());
    }

    private void validate(T r) {
        customValidate(r);
            hibernateValidate(r);

/*
        if (customValidate(r)) {
            hibernateValidate(r);
        }
*/
    }

    private boolean customValidate(T r) {
        for (Field field : fields) {
            String name = field.getName();
            String value = null;
            try {
                field.setAccessible(true);
                value = (String) field.get(r);
                field.setAccessible(false);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
            if (value != null && !"".equals(value) && !notDuplicateValidator.validate(field, name, value, r)) {
                return false;
            }
        }
        return true;
    }

    private void hibernateValidate(T r) {
        Set<ConstraintViolation<T>> validateSet = HibernateValidator.getValidator().validate(r, Default.class);
        if (validateSet != null && !validateSet.isEmpty()) {
            ConstraintViolation<T> constraint = validateSet.stream().findAny().orElse(null);
            r.setError(BindingError.of("email", constraint.getMessage()));
            r.setValidateCode(ExcelRow.FAILED_CODE);
            r.setValidateMessage(constraint.getMessage());
        }
    }

    private void recoverEmptyRow(int target, int real) {
        if (target == real) {
            return;
        }
        List<T> emptyRows = IntStream.rangeClosed(real + headLineMun, target)
                .mapToObj(i -> {
                    try {
                        T t = tClass.newInstance();
                        t.setRowNum(i);
                        hibernateValidate(t);
                        return t;
                    } catch (InstantiationException | IllegalAccessException e) {
                        throw new RuntimeException(e);
                    }
                }).collect(Collectors.toList());
        rows.addAll(emptyRows);
    }
}
