package com.spring.repository.entity;

import com.spring.annotation.Excel;
import lombok.Data;

import java.util.Date;

/**
 * @author Zhendong Zhou
 */
@Data
public class Test {
    // 单元格合并-注意此处的 datum与mergeRow两个属性 datum代表这是一个标识性字段(主键) 工作表数据将会按照该字段的值将 声明了mergeRow的列进行合并行

    @Excel(name = "编码", datum = true)
    private String code;

    @Excel(name = "名称")
    private String name;

    @Excel(name = "时间", mergeRow = true)
    private Date date;
}
