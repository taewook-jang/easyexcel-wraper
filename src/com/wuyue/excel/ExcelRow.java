package com.wuyue.excel;

import com.alibaba.excel.metadata.BaseRowModel;
import lombok.*;

import java.util.ArrayList;
import java.util.List;

/**
 * 基础“行”对象
 */
@Data
public class ExcelRow extends BaseRowModel {

    public static final int SUCCESS_CODE = 0;

    public static final int FAILED_CODE = 2;

    private int rowNum = SUCCESS_CODE;

    private int colNum = SUCCESS_CODE;

    private List<BindingError> bindingErrorList = new ArrayList<>();

    private int validateCode;

    private String validateMessage;

    public boolean hasErrors() {
        return bindingErrorList.size() == 0 ? false : true;
    }

    public void setError(BindingError error) {
        bindingErrorList.add(error);
    }

    @Data
    @RequiredArgsConstructor(staticName = "of")
    public static class BindingError {

        @NonNull
        private String fieldName;
        @NonNull
        private String errorMessage;

    }

}
