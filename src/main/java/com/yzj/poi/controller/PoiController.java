package com.yzj.poi.controller;

import com.yzj.poi.service.impl.PoiService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * 作者: yzj
 * 日期: 2019/10/5
 */
@Controller
@RequestMapping("/stu")
public class PoiController {

    @Autowired
    PoiService poiService;

    @RequestMapping("/export")
    public void exportStu(HttpServletResponse response){

        //设置默认的下载文件名
        String name = "学生信息表.xlsx";
        try {
            //避免文件名中文乱码，将UTF8打散重组成ISO-8859-1编码方式
            name = new String (name.getBytes("UTF8"),"ISO-8859-1");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        //设置响应头的类型
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //让浏览器下载文件,name是上述默认文件下载名
        response.addHeader("Content-Disposition","attachment;filename=\"" + name + "\"");
        InputStream inputStream=null;
        OutputStream outputStream=null;
        //在service层中已经将数据存成了excel临时文件，并返回了临时文件的路径
        String downloadPath = poiService.exportStu();
        //根据临时文件的路径创建File对象，FileInputStream读取时需要使用
        File file = new File(downloadPath);
        try {
            //通过FileInputStream读临时文件，ServletOutputStream将临时文件写给浏览器
            inputStream = new FileInputStream(file);
            outputStream = response.getOutputStream();
            int len = -1;
            byte[] b = new byte[1024];
            while((len = inputStream.read(b)) != -1){
                outputStream.write(b);
            }
            //刷新
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //关闭输入输出流
            try {
                if(inputStream != null) {
                    inputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                if(outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        //最后才能，删除临时文件，如果流在使用临时文件，file.delete()是删除不了的
        file.delete();
    }

}
