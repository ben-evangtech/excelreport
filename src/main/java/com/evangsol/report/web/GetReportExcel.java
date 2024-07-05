package com.evangsol.report.web;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.util.MapUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.fastjson2.JSON;
import com.evangsol.report.fill.MountExcelData;
import com.evangsol.report.util.CustomStyleHandler;
import com.evangsol.report.util.TestFileUtil;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateSerializer;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalTimeSerializer;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.*;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * web读写案例
 *
 **/
@Controller
public class GetReportExcel {



    Logger logger = LoggerFactory.getLogger(GetReportExcel.class);

    /**
     * 文件下载（失败了会返回一个有部分数据的Excel）
     * <p>
     * 1. 创建excel对应的实体对象
     * <p>
     * 2. 设置返回的 参数
     * <p>
     * 3. 直接写，这里注意，finish的时候会自动关闭OutputStream,当然你外面再关闭流问题不大
     */
    @GetMapping("getMountExcel")
    @ResponseBody
    public String getMountExcel(HttpServletResponse response) throws IOException {
        logger.debug("getMountExcel: start");

        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("販売費及び一般管理費", "UTF-8").replaceAll("\\+", "%20");
            logger.debug("getMountExcel: fileName"+fileName);
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            logger.debug("getMountExcel: response.setHeader");


            String templateFileName =
                    TestFileUtil.getPath() + "templates" + File.separator + "販売費及び一般管理費_template.xlsx";
            logger.debug("getMountExcel: templateFileName"+templateFileName);

            MountExcelData data = new MountExcelData();
            data.setDateFrom("2024年3月21日");
            data.setDateTo("2025年3月20日");
            data.setMoment1(1231231231231L);
            data.setMoment11(1150000L);
            data.setMoment25(197000L);
            data.setMoment50(60000L);
            logger.debug("getMountExcel: MountExcelData ready");

//        EasyExcel.write(response.getOutputStream(), MountExcelData.class)
//                .withTemplate(templateFileName)
//                .registerWriteHandler(new CustomStyleHandler())
//                .sheet().doFill(data);

//        FillConfig fillConfig = FillConfig.builder().autoStyle(Boolean.TRUE).build();
//        EasyExcel.write(response.getOutputStream(), MountExcelData.class).withTemplate(templateFileName).sheet().doFill(data);

//        try (ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), MountExcelData.class)
//                .withTemplate(templateFileName)
//                .registerWriteHandler(new CustomStyleHandler())
//                .build()) {
//
//            WriteSheet writeSheet = EasyExcel.writerSheet().build();
//            excelWriter.fill(data, fillConfig, writeSheet);
//        }


            String fileName2 = TestFileUtil.getPath()+fileName+".xlsx";

            logger.debug("getMountExcel: fileName2:"+fileName2);
            EasyExcel.write(fileName2).withTemplate(templateFileName).sheet().doFill(data);
            logger.debug("getMountExcel: EasyExcel end");
            File file = new File(fileName2);
            logger.debug("getMountExcel: file:"+file.getName());


            // 设置响应头
            response.setContentType("application/octet-stream");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            response.setContentLength((int) file.length());
            logger.debug("getMountExcel: response.setContentLength:"+((int) file.length()));

            // 将文件写入响应输出流
            try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
                 BufferedOutputStream bos = new BufferedOutputStream(response.getOutputStream())) {

                logger.debug("getMountExcel: read start");
                byte[] buffer = new byte[1024];
                int bytesRead;

                while ((bytesRead = bis.read(buffer)) != -1) {
                    bos.write(buffer, 0, bytesRead);
                }
                bos.flush();
                logger.debug("getMountExcel: read end");
            } catch (IOException e) {
                throw e;
            }

            logger.debug("getMountExcel: success");
            return "getMountExcel: success";
        } catch (Exception e) {
            e.printStackTrace();
            logger.debug("getMountExcel: error:"+e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            return "getMountExcel: error:"+e.getMessage();
        }
    }



    /**
     * 文件下载（失败了会返回一个有部分数据的Excel）
     * <p>
     * 1. 创建excel对应的实体对象
     * <p>
     * 2. 设置返回的 参数
     * <p>
     * 3. 直接写，这里注意，finish的时候会自动关闭OutputStream,当然你外面再关闭流问题不大
     */
    @PostMapping("postMountExcel")
    @ResponseBody
    public String postMountExcel(HttpServletResponse response, @RequestBody MountExcelData data) throws IOException {
        logger.debug("postMountExcel: start");

        try {
            if (data == null || data.getDateFrom() == null || data.getDateTo() == null) {

                logger.debug("postMountExcel: MountExcelData null");
                return "postMountExcel empty";
            }

            ObjectMapper om = new ObjectMapper();
            JavaTimeModule javaTimeModule = new JavaTimeModule();
            javaTimeModule.addSerializer(LocalDateTime.class, new LocalDateTimeSerializer(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
            javaTimeModule.addSerializer(LocalDate.class, new LocalDateSerializer(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
            javaTimeModule.addSerializer(LocalTime.class, new LocalTimeSerializer(DateTimeFormatter.ofPattern("HH:mm:ss")));
            om.registerModule(javaTimeModule);
            logger.debug("postMountExcel: MountExcelData ready:"+om.writeValueAsString(data));

            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("販売費及び一般管理費", "UTF-8").replaceAll("\\+", "%20");
            logger.debug("postMountExcel: fileName"+fileName);


            String templateFileName =
                    TestFileUtil.getPath() + "templates" + File.separator + "販売費及び一般管理費_template.xlsx";
            logger.debug("postMountExcel: templateFileName"+templateFileName);

            String fileName2 = TestFileUtil.getPath()+fileName+".xlsx";

            logger.debug("postMountExcel: fileName2:"+fileName2);
            EasyExcel.write(fileName2).withTemplate(templateFileName).sheet().doFill(data);
            logger.debug("postMountExcel: EasyExcel end");
//            File file = new File(fileName2);
//            logger.debug("postMountExcel: file:"+file.getName());


            // 设置响应头
//            response.setContentType("application/octet-stream");
//            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
//            logger.debug("postMountExcel: response.setHeader");
//            response.setContentLength((int) file.length());
//            logger.debug("postMountExcel: response.setContentLength:"+((int) file.length()));
//
//            // 将文件写入响应输出流
//            try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
//                 BufferedOutputStream bos = new BufferedOutputStream(response.getOutputStream())) {
//
//                logger.debug("postMountExcel: read start");
//                byte[] buffer = new byte[1024];
//                int bytesRead;
//
//                while ((bytesRead = bis.read(buffer)) != -1) {
//                    bos.write(buffer, 0, bytesRead);
//                }
//                bos.flush();
//                logger.debug("postMountExcel: read end");
//            } catch (IOException e) {
//                throw e;
//            }

            logger.debug("postMountExcel: success");
            return fileName+".xlsx";
        } catch (Exception e) {
            e.printStackTrace();
            logger.debug("postMountExcel: error:"+e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            return "postMountExcel: error:"+e.getMessage();
        }
    }


    /**
     * 文件下载（失败了会返回一个有部分数据的Excel）
     * <p>
     * 1. 创建excel对应的实体对象
     * <p>
     * 2. 设置返回的 参数
     * <p>
     * 3. 直接写，这里注意，finish的时候会自动关闭OutputStream,当然你外面再关闭流问题不大
     */
    @GetMapping("downloadMountExcel")
    public void downloadMountExcel(HttpServletResponse response, String fileNameInput) throws IOException {
        logger.debug("downloadMountExcel: start");

        try {
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode(fileNameInput, "UTF-8").replaceAll("\\+", "%20");
            logger.debug("downloadMountExcel: fileName"+fileName);
            response.setHeader("Content-disposition", "attachment;filename=" + fileName);
            logger.debug("downloadMountExcel: response.setHeader");


            String fileNameDownload = TestFileUtil.getPath()+fileName;

            logger.debug("downloadMountExcel: fileNameDownload:"+fileNameDownload);
            File file = new File(fileNameDownload);
            logger.debug("downloadMountExcel: file:"+file.getName());


            // 设置响应头
            response.setContentType("application/octet-stream");
            response.setContentLength((int) file.length());
            logger.debug("downloadMountExcel: response.setContentLength:"+((int) file.length()));

            // 将文件写入响应输出流
            try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
                 BufferedOutputStream bos = new BufferedOutputStream(response.getOutputStream())) {

                logger.debug("downloadMountExcel: read start");
                byte[] buffer = new byte[1024];
                int bytesRead;

                while ((bytesRead = bis.read(buffer)) != -1) {
                    bos.write(buffer, 0, bytesRead);
                }
                bos.flush();
                logger.debug("downloadMountExcel: read end");
            } catch (IOException e) {
                throw e;
            }

            logger.debug("downloadMountExcel: success");
        } catch (Exception e) {
            e.printStackTrace();
            logger.debug("downloadMountExcel: error:"+e.getMessage());
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
    }

    @GetMapping("testApi")
    @ResponseBody
    public String testApi()  {
        return "success";
    }


}
