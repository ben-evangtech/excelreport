package com.evangsol.report.web;

import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.util.MapUtils;

import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.fastjson2.JSON;
import com.evangsol.report.fill.FillData;
import com.evangsol.report.fill.MountExcelData;
import com.evangsol.report.util.TestFileUtil;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

/**
 * web读写案例
 *
 **/
@Controller
public class WebTest {


    /**
     * 文件下载（失败了会返回一个有部分数据的Excel）
     * <p>
     * 1. 创建excel对应的实体对象
     * <p>
     * 2. 设置返回的 参数
     * <p>
     * 3. 直接写，这里注意，finish的时候会自动关闭OutputStream,当然你外面再关闭流问题不大
     */
    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("TEST", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

//        EasyExcel.write(response.getOutputStream(), FillData.class).sheet("模板").doWrite(data());

//        String templateFileName =
//                TestFileUtil.getPath() + "templates" + File.separator + "simple.xlsx";
        String templateFileName =
                TestFileUtil.getPath() + "templates" + File.separator + "mount_template_en.xlsx";
        MountExcelData data = new MountExcelData();
        data.setDateFrom("2024/03/21");
        data.setDateTo("2025/03/20");
        data.setMoment1(1231231231231L);
        data.setMoment11(1150000L);
        data.setMoment25(197000L);
        data.setMoment50(60000L);
        // 方案1 根据对象填充
//        String fileName = TestFileUtil.getPath() + "simpleFill" + System.currentTimeMillis() + ".xlsx";

//        EasyExcel.write(response.getOutputStream(), FillData.class).sheet("模板").doWrite(data());
//        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(fillData);
//        EasyExcel.write(response.getOutputStream(), FillData.class).withTemplate(templateFileName).sheet().doFill(data());
        try (ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), MountExcelData.class).withTemplate(templateFileName).build()) {
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            excelWriter.fill(data, writeSheet);
//            excelWriter.fill(data(), writeSheet);
        }
    }

    /**
     * 文件下载并且失败的时候返回json（默认失败了会返回一个有部分数据的Excel）
     *
     * @since 2.1.1
     */
    @GetMapping("downloadFailedUsingJson")
    public void downloadFailedUsingJson(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream(), FillData.class).autoCloseStream(Boolean.FALSE).sheet("模板")
                .doWrite(data());
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
//            Map<String, String> map = MapUtils.newHashMap();
//            map.put("status", "failure");
//            map.put("message", "下载文件失败" + e.getMessage());
//            response.getWriter().println(JSON.toJSONString(map));
        }
    }



    private List<FillData> data() {
        List<FillData> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            list.add(fillData);
            fillData.setName("张三");
            fillData.setNumber(5.2);
            fillData.setDate(new Date());
        }
        return list;
    }
}
