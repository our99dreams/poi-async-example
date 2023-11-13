package com.spring.repository.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.spring.repository.entity.Test;
import org.apache.ibatis.annotations.Options;
import org.apache.ibatis.cursor.Cursor;
import org.apache.ibatis.mapping.ResultSetType;

/**
 * @author Zhendong Zhou
 */
public interface TestMapper extends BaseMapper<Test> {
    @Options(resultSetType = ResultSetType.FORWARD_ONLY, fetchSize = Integer.MIN_VALUE)
    Cursor<Test> cursor();
}
