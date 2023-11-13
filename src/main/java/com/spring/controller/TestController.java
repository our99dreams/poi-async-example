package com.spring.controller;

import com.spring.service.TestService;
import com.spring.util.ExcelExportUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Map;

/**
 * @author Zhendong Zhou
 * @since 2023-11-12
 */
@Slf4j
@RestController
@RequestMapping("test")
public class TestController {
    @Autowired
    public TestController(TestService service,
                          RedisTemplate redisTemplate) {
        this.service = service;
        this.redisTemplate = redisTemplate;
    }
    private final TestService service;
    private final RedisTemplate<String, Object> redisTemplate;

    // 待办的任务列表
    @GetMapping(value = "task-list")
    public Map<Object, Object> task() {
        return ExcelExportUtils.ExcelTaskUtil.getExportTask(redisTemplate, 1L);
    }

    // 下载文件
    @GetMapping(value = "download")
    public void download(@RequestParam("fileName") String filename, HttpServletRequest request, HttpServletResponse response) {
        if (!ExcelExportUtils.ExcelTaskUtil.verifyTaskExist(redisTemplate, filename, 1L)) {
            throw new RuntimeException("任务已过期");
        }
        String filePath = ExcelExportUtils.PUBLIC_EXPORT_DIR + filename;
        if (!new File(filePath).exists()) {
            ExcelExportUtils.ExcelTaskUtil.removeCache(redisTemplate, filename, 1L);
            throw new RuntimeException("文件已被清理");
        }
        try {
            response.setHeader("requestType","file");
            response.setHeader("Access-Control-Expose-Headers","requestType");
            String filePrefixName = getFilePrefixName(filename);
            compatibleFileName(request, response, filePrefixName);
            writeBytes(filePath, response.getOutputStream());
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }
    }

    private static String getFilePrefixName(String filename){
        if (StringUtils.isEmpty(filename) && filename.indexOf("\\.")>0){
            return "";
        }
        String[] name = filename.split("\\.");
        return name[0];
    }

    private static void compatibleFileName(HttpServletRequest request, HttpServletResponse response, String excelName)
            throws UnsupportedEncodingException {
        String agent = request.getHeader("USER-AGENT").toLowerCase();
        response.setContentType("application/vnd.ms-excel");
        String codedFileName = java.net.URLEncoder.encode(excelName, StandardCharsets.UTF_8.name());
        if (agent.contains("firefox") || agent.contains("safari")) {
            response.setCharacterEncoding(StandardCharsets.UTF_8.name());
            response.setHeader("content-disposition", "attachment;filename=" + new String(excelName.getBytes(), "ISO8859-1") + ".xlsx");
        } else {
            response.setHeader("content-disposition", "attachment;filename=" + codedFileName + ".xlsx");
        }
    }

    private static void writeBytes(String filePath, OutputStream os) throws IOException {
        FileInputStream fis = null;
        try {
            File file = new File(filePath);
            if (!file.exists()) {
                throw new FileNotFoundException(filePath);
            }
            fis = new FileInputStream(file);
            byte[] b = new byte[1024];
            int length;
            while ((length = fis.read(b)) > 0) {
                os.write(b, 0, length);
            }
        } catch (IOException e) {
            throw e;
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }
    }
}
