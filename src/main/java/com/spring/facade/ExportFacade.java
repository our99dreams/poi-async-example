package com.spring.facade;

import com.alibaba.fastjson.JSONObject;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import com.spring.util.ExcelExportUtils;
import org.apache.ibatis.cursor.Cursor;
import org.springframework.data.redis.core.RedisTemplate;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Lazy;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;
import org.springframework.stereotype.Component;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.transaction.support.TransactionTemplate;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

/**
 * 导出门面
 *
 * @author Zhendong Zhou
 * @date 2023/6/7
 */
@Slf4j
@Component
@RequiredArgsConstructor(onConstructor_ = {@Lazy, @Autowired})
public class ExportFacade {
    private final RedisTemplate<String, Object> redisService;
    private final ThreadPoolTaskExecutor poolTaskExecutor;
    private final PlatformTransactionManager dataSourceTransactionManager;

    private int pages = 1; // 初始总页数
    private int size = ExcelExportUtils.MAX_EXPORT_COUNT; // 初始条数
    public ExportFacade setPages(int pages) { this.pages = pages; return this; }
    public ExportFacade setSize(int size) { this.size = size; return this; }

    private <T> void autoPages(DataProvider<T> dataService) {
        Page<T> page = new Page<>(0, 0);
        page.setSearchCount(true);
        long total = dataService.page(page).getTotal();
        this.pages = (int)Math.ceil(((double)total) / size);
    }
    public <T> void execute(String fileName, Class<T> clazz, DataProvider<T> dataService, long userId) {
        this.execute(fileName, clazz, dataService, false, userId);
    }

    public <T> void execute(String fileName, Class<T> clazz, DataProvider<T> dataService, boolean auto, long userId) {
        ExcelExportUtils<T> exportUtil = new ExcelExportUtils<>(clazz,
                ExcelExportUtils.PUBLIC_EXPORT_DIR, fileName);
        ExcelExportUtils.ExcelTaskUtil.addExportFile(redisService, exportUtil.getFileName(), userId, pages);
        if (auto) {
            this.autoPages(dataService);
        }

        for (int i = 1; i <= pages; i++) {
            Page<T> page = new Page<>(i, size);
            page.setSearchCount(false);
            poolTaskExecutor.execute(() -> {
                try{
                    List<T> list = dataService.page(page).getRecords();
                    exportUtil.build(list, exportUtil.buildSheet((int)(page.getCurrent() - 1)));
                } catch (Exception e) {
                    log.error("线程池异常：{},{}", e.getMessage(),
                            JSONObject.toJSONString(e.getStackTrace()));
                } finally {
                    if (ExcelExportUtils.ExcelTaskUtil.scheduleExportFile(redisService, exportUtil.getFileName(), userId)) {
                        exportUtil.make();
                    }
                }
            });
        }
    }

    public ExportFacade setCursor(int handlerSize) {
        this.handlerSize = handlerSize;
        return this;
    }
    private int handlerSize = 1000;

    /**
     * 游标处理-游标只有在事务中才会生效
     *
     * @param fileName 文件名
     * @param clazz 目标类
     * @param dataService 数据供给
     * @param handler 额外的分批处理器-游标数据每N条进行一次截片处理（可以用以拼接第三方数据）-可以使用setCursor来指定分片大小
     * @param userId 用户ID
     * @param <T> 目标对象-存在属性必须标注了Excel注解
     */
    public <T> void cursor(String fileName, Class<T> clazz, CursorProvider<T> dataService, Consumer<List<T>> handler, long userId) {
        // 创建一个事务模板用默认的数据源事务管理器
        TransactionTemplate transactionTemplate = new TransactionTemplate(dataSourceTransactionManager);
        ExcelExportUtils<T> exportUtil = new ExcelExportUtils<>(ExcelExportUtils.PUBLIC_EXPORT_DIR, fileName);

        AtomicInteger i = new AtomicInteger();
        poolTaskExecutor.execute(() -> transactionTemplate.execute(action -> {
            Cursor<T> cursor = null;
            try {
                // 获取页写入对象
                AtomicReference<ExcelExportUtils<?>.Write> atomic = new AtomicReference<>();
                ExcelExportUtils<?>.Write write = exportUtil.build(exportUtil.buildSheet("第" + (i.get() + 1) + "页"),
                        exportUtil.buildExcelField(clazz));
                atomic.set(write);

                cursor = dataService.cursor();
                // 每收集一定条目进行一次处理
                Consumer<List<T>> resultHandler = (data) -> {
                    if (handler != null) {
                        // TODO 对每批数据做一些额外的处理
                        handler.accept(data);
                    }
                    if (atomic.get().size() + data.size() > size) {
                        atomic.get().close();
                        i.getAndIncrement();
                        atomic.set(exportUtil.build(exportUtil.buildSheet("第" + (i.get() + 1) + "页"),
                                exportUtil.buildExcelField(clazz)));
                    }
                    atomic.get().append(data);
                };
                // 分批处理游标数据，每收集一部分进行处理,这样是会多几次循环操作 但是可以保证职责单一与良好的隔离性
                this.handler(cursor, resultHandler, handlerSize);
                write.close();
            } catch (Exception e) {
                log.error("数据导出失败：{},堆栈信息：{}", e.getMessage(), JSONObject.toJSON(e.getStackTrace()));
            } finally {
                try {
                    if (cursor != null) { cursor.close(); }
                } catch (IOException ignored) {}
                if (ExcelExportUtils.ExcelTaskUtil.scheduleExportFile(redisService, exportUtil.getFileName(), userId)) {
                    exportUtil.make();
                }
            }
            return null;
        }));
    }

    private <T> void handler(Cursor<T> cursor, Consumer<List<T>> handler, int size) {
        List<T> batch = new ArrayList<>(size);
        for (T t : cursor) {
            batch.add(t);
            if (batch.size() == size) {
                handler.accept(batch);
                batch.clear();
            }
        }
        if (!batch.isEmpty()) handler.accept(batch);
    }

    public interface DataProvider<T> {
        Page<T> page(Page<T> page);
    }
    public interface CursorProvider<T> {
        Cursor<T> cursor();
    }
}
