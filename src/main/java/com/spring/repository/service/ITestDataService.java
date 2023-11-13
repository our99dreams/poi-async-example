package com.spring.repository.service;

import com.baomidou.mybatisplus.extension.service.IService;
import com.spring.repository.entity.Test;
import org.apache.ibatis.cursor.Cursor;

/**
 * @author Zhendong Zhou
 */
public interface ITestDataService extends IService<Test> {
    Cursor<Test> cursor();
}
