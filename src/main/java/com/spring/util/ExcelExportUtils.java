package com.spring.util;

import com.alibaba.fastjson.JSONObject;
import com.spring.annotation.Excel;
import com.spring.annotation.Excels;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;

/**
 * 实验多线程导出工具包
 *
 * @author Zhendong Zhou
 * @date 2022/7/8
 */
@Slf4j
public class ExcelExportUtils<T> {
    public final static int MAX_EXPORT_COUNT = 50000;// 单页最大导出数量
    public final static String PUBLIC_EXPORT_DIR = "/share/storage/export/";// 默认的导出的目录


    /**
     * 实体对象
     */
    private Class<T> clazz;
    /**
     * 注解列表
     */
    private List<Field> fields;

    /**
     * 文件对象
     */
    private SXSSFWorkbook book;

    /**
     * 文件名
     */
    private final String fileName;

    private final String filePath;

    public ExcelExportUtils(Class<T> clazz, String dir, String moduleName) {
        this(dir, moduleName);

        this.clazz = clazz;
        this.buildExcelField();
    }
    public ExcelExportUtils(String dir, String moduleName) {
        this.buildWorkBook();
        this.fileName = encodingFilename(moduleName);
        this.filePath = getAbsoluteFile(dir, fileName);
    }
    /**
     * 编码文件名
     */
    public String encodingFilename(String filename) {
        filename = filename + "-" + System.currentTimeMillis() + ".xlsx";
        return filename;
    }

    /**
     * 返回预期文件名
     *
     * @return 文件名
     */
    public String getFileName() {
        return fileName;
    }
    /**
     * 获取下载路径
     *
     * @param filename 文件名称
     */
    public String getAbsoluteFile(String filePath, String filename) {
        String downloadPath = filePath + filename;
        File desc = new File(downloadPath);
        if (!desc.getParentFile().exists()) {
            if (!desc.getParentFile().mkdirs()) {
                log.error("建立存储Excel目录失败");
            }
        }
        return downloadPath;
    }

    public void buildWorkBook() {
        this.book = new SXSSFWorkbook(MAX_EXPORT_COUNT);
    }

    public SXSSFSheet buildSheet(int start) {
        return book.createSheet(((start * MAX_EXPORT_COUNT)+1)+"-"+((start+1)* MAX_EXPORT_COUNT));
    }
    public synchronized SXSSFSheet buildSheet(String name) {
        return book.createSheet(name);
    }

    public String make() {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            book.write(out);
        } catch (Exception e) {
            log.error("EXCEL文件生成失败:{}",e.getMessage());
            throw new RuntimeException("EXCEL文件生成失败");
        } finally {
            if (book != null) {
                try {
                    book.close();
                } catch (IOException e1) {
                    log.error("EXCEL对象资源关闭失败:{}",fileName);
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e1) {
                    log.error("EXCEL文件资源关闭失败:{}",fileName);
                }
            }
        }
        return fileName;
    }
    public void build(List<T> list, SXSSFSheet sheet) {
        this.build(list, sheet, fields);
    }
    private void build(List<T> list, SXSSFSheet sheet, List<Field> fields) {
        Write write = this.build(sheet, fields);
        write.append(list);
        write.close();
    }
    public Write build(SXSSFSheet sheet, List<Field> fields) {
        // 标题行
        Row row = sheet.createRow(0);
        int excelsNo = 0;
        // 写入各个字段的列头名称
        for (int column = 0; column < fields.size(); column++) {
            Field field = fields.get(column);
            if (field.isAnnotationPresent(Excel.class)) {
                Excel excel = field.getAnnotation(Excel.class);
                createCell(sheet, excel, row, column);
            }
            if (field.isAnnotationPresent(Excels.class)) {
                Excels attrs = field.getAnnotation(Excels.class);
                Excel[] excels = attrs.value();
                // 写入列名
                Excel excel = excels[excelsNo++];
                createCell(sheet, excel, row, column);
            }
        }
        CellStyle cs = book.createCellStyle();
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        Excel[] excels = new Excel[fields.size()];
        Integer datum = null;
        List<Integer> mergeColumn = new ArrayList<>();
        for (int column = 0; column < fields.size(); column++) {
            // 获得field.
            Field field = fields.get(column);
            // 设置实体类私有属性可访问
            field.setAccessible(true);
            if (field.isAnnotationPresent(Excel.class)) {
                excels[column] = field.getAnnotation(Excel.class);
                if (excels[column].datum()) {
                    datum = column;
                }
                if (excels[column].mergeRow()) {
                    mergeColumn.add(column);
                }
            }
        }
        SlowSign slow = new SlowSign(mergeColumn, sheet);

        return new Write(sheet, datum, excels, slow, cs, fields);
    }
    public class Write {
        private final Sheet sheet;
        private final Integer datum;
        private final Excel[] excels;
        private final SlowSign slow;
        private final CellStyle cs;
        private final List<Field> fields;

        private int rowNum = 1;
        public Write(Sheet sheet, Integer datum, Excel[] excels, SlowSign slow, CellStyle cs, List<Field> fields) {
            this.sheet = sheet;
            this.datum = datum;
            this.excels = excels;
            this.slow = slow;
            this.cs = cs;
            this.fields = fields;
        }
        public int size() {
            return rowNum - 1;
        }

        public void append(Object rowData) {
            Row row = sheet.createRow(rowNum);
            // 是否需要合并
            if (datum != null) {
                try{
                    String val = String.valueOf(getTargetValue((T)rowData, fields.get(datum), excels[datum]));
                    slow.merge(val, rowNum);
                } catch (Exception e) {
                    log.error("导出Excel失败{},{}", e.getMessage(), JSONObject.toJSONString(e.getStackTrace()));
                    return;
                }
            }
            for (int column = 0; column < fields.size(); column++) {
                if (excels[column] == null) {
                    continue;
                }
                addCell(excels[column], row, (T)rowData, fields.get(column), column, cs);
            }
            rowNum++;
        }
        public void append(List<?> data) {
            for (Object rowData : data) { this.append(rowData); }
        }

        public void close() {
            if (datum != null) { slow.endMerge(); }
        }
    }

    public List<Field> buildExcelField(Class<?> clazz) {
        ArrayList<Field> fields = new ArrayList<>();
        List<Field> tempFields = new ArrayList<>();
        tempFields.addAll(Arrays.asList(clazz.getSuperclass().getDeclaredFields()));
        tempFields.addAll(Arrays.asList(clazz.getDeclaredFields()));
        for (Field field : tempFields) {
            // 单注解
            if (field.isAnnotationPresent(Excel.class)) {
                Excel attr = field.getAnnotation(Excel.class);
                if (attr != null && (attr.type() == Excel.Type.ALL || attr.type() == Excel.Type.EXPORT)) {
                    fields.add(field);
                }
            }

            // 多注解
            if (field.isAnnotationPresent(Excels.class)) {
                Excels attrs = field.getAnnotation(Excels.class);
                Excel[] excels = attrs.value();
                for (Excel excel : excels) {
                    if (excel != null && (excel.type() == Excel.Type.ALL || excel.type() == Excel.Type.EXPORT)) {
                        fields.add(field);
                    }
                }
            }
        }
        return fields;
    }

    private void buildExcelField() {
        this.fields = this.buildExcelField(clazz);
    }

    /**
     * 添加单元格
     */
    public void addCell(Excel attr, Row row, T vo, Field field, int column, CellStyle cs) {
        Cell cell;
        try {
            // 设置行高
            row.setHeight((short) (attr.height() * 20));
            // 根据Excel中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
            if (attr.isExport()) {
                // 创建cell
                cell = row.createCell(column);
                cell.setCellStyle(cs);

                // 用于读取对象中的属性
                Object value = getTargetValue(vo, field, attr);
                String dateFormat = attr.dateFormat();
                boolean isDecimalFormat = attr.isDecimalFormat();
                if (value instanceof Date) {
                    cell.setCellValue(new SimpleDateFormat(dateFormat).format((Date)value));
                } else if (isDecimalFormat) {
                    if (value == null) {
                        cell.setCellValue(0);
                    } else {
                        cell.setCellValue((new BigDecimal(value.toString())).stripTrailingZeros().toPlainString());
                    }
                } else {
                    cell.setCellType(attr.cellType());
                    if (value == null) {
                        cell.setCellValue(attr.defaultValue());
                    } else {
                        switch (attr.cellType()) {
                            case NUMERIC:
                                cell.setCellValue(Double.parseDouble(String.valueOf(value)));
                                break;
                            default:
                                cell.setCellValue(value + attr.suffix());
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("导出Excel失败{},{}", e.getMessage(), JSONObject.toJSONString(e.getStackTrace()));
        }
    }

    /**
     * 获取bean中的属性值
     *
     * @param vo    实体对象
     * @param field 字段
     * @param excel 注解
     * @return 最终的属性值
     */
    private Object getTargetValue(T vo, Field field, Excel excel) throws Exception {
        Object o = field.get(vo);
        if (!StringUtils.isEmpty(excel.targetAttr())) {
            String target = excel.targetAttr();
            if (target.contains(".")) {
                String[] targets = target.split("[.]");
                for (String name : targets) {
                    o = getValue(o, name);
                }
            } else {
                o = getValue(o, target);
            }
        }
        return o;
    }

    /**
     * 解析导出值 0=男,1=女,2=未知
     *
     * @param propertyValue 参数值
     * @param converterExp  翻译注解
     * @return 解析后值
     */
    public static String convertByExp(String propertyValue, String converterExp) {
        String[] convertSource = converterExp.split(",");
        for (String item : convertSource) {
            String[] itemArray = item.split("=");
            if (itemArray[0].equals(propertyValue)) {
                return itemArray[1];
            }
        }
        return propertyValue;
    }

    /**
     * 以类的属性的get方法方法形式获取值
     *
     * @param o 属性对象
     * @param name 属性名
     * @return value 值
     */
    private Object getValue(Object o, String name) throws Exception {
        if (!StringUtils.isEmpty(name)) {
            Class<?> clazz = o.getClass();
            String methodName = "get" + name.substring(0, 1).toUpperCase() + name.substring(1);
            Method method = clazz.getMethod(methodName);
            o = method.invoke(o);
        }
        return o;
    }

    /**
     * 创建单元格
     */
    public void createCell(Sheet sheet, Excel attr, Row row, int column) {
        // 创建列
        Cell cell = row.createCell(column);
        // 设置列中写入内容为String类型
        cell.setCellType(CellType.STRING);
        // 写入列名
        cell.setCellValue(attr.name());
        CellStyle cellStyle = createStyle(sheet, attr, row, column);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 创建表格样式
     */
    public CellStyle createStyle(Sheet sheet, Excel attr, Row row, int column) {
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = book.createFont();
        if (attr.name().contains("注：")) {
            font.setColor(HSSFFont.COLOR_RED);
            cellStyle.setFont(font);
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
            sheet.setColumnWidth(column, 6000);
        } else {
            // 粗体显示
            font.setBold(true);
            // 选择需要用到的字体格式
            cellStyle.setFont(font);
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
            // 设置列宽
            sheet.setColumnWidth(column, (int) ((attr.width() + 0.72) * 256));
            row.setHeight((short) (attr.height() * 20));
        }
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setWrapText(true);
        // 如果设置了提示信息则鼠标放上去提示.
        if (!StringUtils.isEmpty(attr.prompt())) {
            // 这里默认设了2-101列提示.
            setXSSFPrompt(sheet, "", attr.prompt(), 1, 100, column, column);
        }
        // 如果设置了combo属性则本列只能选择不能输入
        if (attr.combo().length > 0) {
            // 这里默认设了2-101列只能选择不能输入.
            setXSSFValidation(sheet, attr.combo(), 1, 100, column, column);
        }
        return cellStyle;
    }

    /**
     * 设置 POI XSSFSheet 单元格提示
     *
     * @param sheet         表单
     * @param promptTitle   提示标题
     * @param promptContent 提示内容
     * @param firstRow      开始行
     * @param endRow        结束行
     * @param firstCol      开始列
     * @param endCol        结束列
     */
    public void setXSSFPrompt(Sheet sheet, String promptTitle, String promptContent, int firstRow, int endRow,
                              int firstCol, int endCol) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createCustomConstraint("DD1");
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        DataValidation dataValidation = helper.createValidation(constraint, regions);
        dataValidation.createPromptBox(promptTitle, promptContent);
        dataValidation.setShowPromptBox(true);
        sheet.addValidationData(dataValidation);
    }

    /**
     * 设置某些列的值只能输入预制的数据,显示下拉框.
     *
     * @param sheet    要设置的sheet.
     * @param textlist 下拉框显示的内容
     * @param firstRow 开始行
     * @param endRow   结束行
     * @param firstCol 开始列
     * @param endCol   结束列
     *
     */
    public void setXSSFValidation(Sheet sheet, String[] textlist, int firstRow, int endRow, int firstCol, int endCol) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        // 加载下拉列表内容
        DataValidationConstraint constraint = helper.createExplicitListConstraint(textlist);
        // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        // 数据有效性对象
        DataValidation dataValidation = helper.createValidation(constraint, regions);
        // 处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }

        sheet.addValidationData(dataValidation);
    }

    static class SlowSign {
        int startRow;
        String lastValue;
        int lastRow;
        List<Integer> columns;
        SXSSFSheet sheet;
        SlowSign(List<Integer> columns, SXSSFSheet sheet) {
            this.lastValue = "";
            this.columns = columns;
            this.sheet = sheet;
        }
        public void merge(String value, int row) {
            if (!lastValue.equals(value)) {
                if (startRow != lastRow) {
                    this.setMerge();
                }
                lastValue = value == null ? "" : value;
                startRow = row;
            }
            this.lastRow = row;
        }

        public void endMerge() {
            if (startRow != lastRow) {
                this.setMerge();
            }
        }

        private void setMerge() {
            for (Integer column : columns) {
                CellRangeAddress mergeAddress = new CellRangeAddress(startRow, lastRow, column, column);
                sheet.addMergedRegionUnsafe(mergeAddress);
                // 删除被合并cell,暂时不知道为什么设置addMergedRegionUnsafe不会清除所属cell，所以需要手动删除一下
                for (int i = startRow+1; i <= lastRow; i++) {
                    sheet.getRow(i).removeCell(sheet.getRow(i).getCell(column));
                }
            }
        }
    }


    /**
     *  Excel相关的任务缓存工具类
     */
    public static class ExcelTaskUtil {
        private static long EXPIRE = 86500;// 过期时间

        /**
         * REDIS缓存KEY值
         *
         * @param id 用户ID
         * @return KEY值
         */
        static String USER_EXPORT_TASK(Long id) {
            return "user:export:task:".concat(String.valueOf(id));
        }

        /**
         * 添加一个文件缓存任务
         *
         * @param redisTemplate 缓存服务
         * @param fileName 文件名
         * @param id 用户ID
         * @param countGroup 最大执行总量组
         */
        public static void addExportFile(RedisTemplate<String, Object> redisTemplate, String fileName, Long id, int countGroup) {
            String key = USER_EXPORT_TASK(id);
            if (!redisTemplate.hasKey(key)) {
                redisTemplate.opsForHash().putAll(key, new HashMap<>(16));
            }
            redisTemplate.opsForHash().put(key, fileName, "0/".concat(String.valueOf(countGroup)));
            // 文件一天过期，缓存信息一天过期 不过期也会被系统清理
            redisTemplate.expire(key, EXPIRE, TimeUnit.SECONDS);
        }

        /**
         * 执行计划导出文件/片段执行结果
         *
         * @param redisTemplate 缓存服务
         * @param fileName 文件名
         * @param id 用户ID
         * @return 任务是否全部执行完成
         */
        public static boolean scheduleExportFile(RedisTemplate<String, Object> redisTemplate, String fileName, Long id) {
            String key = USER_EXPORT_TASK(id);
            String val = (String)redisTemplate.opsForHash().get(key, fileName);
            if (val != null) {
                String[] arr = val.split("/");
                int current = Integer.parseInt(arr[0])+1;
                int count = Integer.parseInt(arr[1]);

                redisTemplate.opsForHash().put(key, fileName, current+"/"+count);
                // 文件一天过期，缓存信息一天过期 不过期也会被系统清理
                redisTemplate.expire(key, EXPIRE, TimeUnit.SECONDS);
                return current >= count;
            }
            return false;
        }

        /**
         * 获取用户下全部导出任务
         *
         * @param redisTemplate 缓存服务
         * @param id 用户ID
         * @return 任务列表
         */
        public static Map<Object, Object> getExportTask(RedisTemplate<String, Object> redisTemplate, Long id) {
            String key = USER_EXPORT_TASK(id);
            if (!redisTemplate.hasKey(key)) {
                return new HashMap<>(0);
            }
            return redisTemplate.opsForHash().entries(key);
        }

        /**
         * 验证任务是否存在
         *
         * @param redisTemplate 缓存服务
         * @param fileName 文件名
         * @param userId 用户ID
         * @return 验证结果
         */
        public static boolean verifyTaskExist(RedisTemplate<String, Object> redisTemplate, String fileName, Long userId) {
            String key = USER_EXPORT_TASK(userId);
            return redisTemplate.hasKey(key) && redisTemplate.opsForHash().hasKey(key, fileName);
        }

        public static void removeCache(RedisTemplate<String, Object> redisTemplate, String fileName, Long userId) {
            String key = USER_EXPORT_TASK(userId);
            redisTemplate.opsForHash().delete(key, fileName);
        }
    }
}
