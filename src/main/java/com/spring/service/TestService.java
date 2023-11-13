package com.spring.service;

import com.spring.facade.ExportFacade;
import com.spring.repository.entity.Test;
import com.spring.repository.service.ITestDataService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Lazy;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.PlatformTransactionManager;

/**
 * @author Zhendong Zhou
 */
@Slf4j
@Service
@RequiredArgsConstructor(onConstructor_ = {@Lazy, @Autowired})
public class TestService {
    private final ExportFacade exportFacade;
    private final ITestDataService testDataService;
    private final ThreadPoolTaskExecutor poolTaskExecutor;
    private final PlatformTransactionManager dataSourceTransactionManager;
    private final RedisTemplate<String, Object> redisTemplate;

    private long userId = 1;

    // 普通用法-自动查询条数分页
    public void commonExport() {
        exportFacade.execute("测试", Test.class, testDataService::page, userId);
    }

    // 手动分页-禁止自动分页，手动指定查询数目与分页条数,默认情况下如果假定数据量不会超过初始值可以直接使用该方法
    public void pageExport() {
        exportFacade.setPages(2).setSize(1000).execute("测试", Test.class, testDataService::page, false, userId);
    }

    // 游标用法
    public void cursorExport() {
        exportFacade.setCursor(2000).cursor("测试", Test.class, testDataService::cursor, items -> {
            for (Test item : items) {
                item.setName(item.getName() + "-游标的额外处理");
            }
        }, userId);
    }
}
