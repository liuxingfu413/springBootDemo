package com.exa.ctl;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.StyleSet;
import com.exa.beans.exportBean;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@RestController
public class testController {

    @Value("${localPath}")
    private String localPath;

    /**
     * http://127.0.0.1:8080/test1
     * @return
     */
    @GetMapping("/test1")
    public String test1(){
        System.out.println("222");
        return "success";

    }

    /**
     * 测试月份字符串获取上个月时间
     * http://127.0.0.1:8080/timeTest
     * @return
     */
    @GetMapping("/timeTest")
    public String timeTest(){
        String inputDate = "2023-01";
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM");
        YearMonth yearMonth = YearMonth.parse(inputDate, formatter);
        YearMonth lastYearMonth = yearMonth.minusMonths(1);
        String lastMonthString = lastYearMonth.format(formatter);

        System.out.println("Last month: " + lastMonthString);
        return lastMonthString;

    }

    @ResponseBody
    @GetMapping("/exportTest")
    public ResponseEntity<byte[]> exportTest() throws IOException {
        List<exportBean> list = getExport();
        // 创建ExcelWriter
        ExcelWriter writer = ExcelUtil.getWriter(true);

        //自定义样式
        StyleSet styleSet = writer.getStyleSet();
        CellStyle cellStyle = styleSet.getHeadCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        //设置列宽
        writer.setColumnWidth(0,16);
        writer.setColumnWidth(1,16);
        writer.setColumnWidth(2,50);

        //设置表头
        writer.addHeaderAlias("id", "序号");
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("text", "内容");

        //写入数据
        writer.setOnlyAlias(true); // 只导出设置了别名的字段
        writer.write(list);

//        writer.write(list, false); 不把字段名作为表头，只导出字段的内容


        // 生成Excel文件到本地
        String fileName = UUID.randomUUID().toString().replace("-","") + ".xls";
        String filePath = localPath + fileName;
        File outFile = new File(filePath);
        // 判断目标文件所在目录知否存在，不存在则创建父目录
        if (!outFile.getParentFile().exists()){
            outFile.getParentFile().mkdir();
        }
        writer.flush(FileUtil.getOutputStream(filePath));

        writer.close();

        //组装报文头下载
        File file = new File(filePath);
        FileInputStream fileInputStream = new FileInputStream(file);
        byte[] bytes = new byte[(int) file.length()];
        fileInputStream.read(bytes);
        fileInputStream.close();

        HttpHeaders headers = new HttpHeaders();
        // 设置响应的内容类型为 MediaType.APPLICATION_OCTET_STREAM，表示下载的是二进制文件
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        // 设置响应头中的 Content-Disposition，指定文件的名称和下载方式。
        // Content-Disposition 是一个标准的 HTTP 响应头字段，用于指示浏览器如何处理接收到的响应内容。它可以有两种下载方式：
        //  attachment：表示将响应内容作为附件下载。浏览器会弹出文件下载对话框，提示用户保存文件到本地。
        //  inline：表示将响应内容直接在浏览器中打开。浏览器会尝试使用适当的插件或内置的查看器来打开文件，例如在浏览器中直接显示 PDF 文件。
        headers.setContentDispositionFormData("attachment", file.getName());


        return new ResponseEntity<byte[]>(bytes, headers, HttpStatus.OK);


    }

    private List<exportBean> getExport(){
        List<exportBean> list = new ArrayList<>();
        for (int i = 0; i < 10; i++){
            exportBean exportBean = new exportBean();
            exportBean.setId(String.valueOf(i));
            exportBean.setName("姓名" + i);
            exportBean.setText("导出内容" + i);
            list.add(exportBean);
        }
        return list;
    }


}
