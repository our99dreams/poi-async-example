package com.spring.repository.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.spring.repository.entity.Test;
import com.spring.repository.mapper.TestMapper;
import com.spring.repository.service.ITestDataService;
import org.apache.ibatis.cursor.Cursor;
import org.springframework.stereotype.Service;

/**
 * @author Zhendong Zhou
 */
@Service
public class TestDataServiceImpl extends ServiceImpl<TestMapper, Test> implements ITestDataService {
    @Override
    public Cursor<Test> cursor() {
        return baseMapper.cursor();
    }
}
