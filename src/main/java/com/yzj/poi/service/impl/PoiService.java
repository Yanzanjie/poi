package com.yzj.poi.service.impl;

import com.yzj.poi.po.Student;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * 作者: yzj
 * 日期: 2019/10/5
 */
@Service
public class PoiService {

    //创建临时文件存放的路径
    private String temp="d:"+ File.separator+"3"+File.separator+"temporary"+File.separator;

//    @Autowired
//    Student student;

    public List<Student> allStu(){
        List<Student> liststu = new ArrayList<Student>();
//        Student allstu = new Student();

        for(int i = 0 ; i<10 ;i++){
            Student allstu = new Student();
            allstu.setId(i);
            allstu.setName("lisi"+i);
            allstu.setAge("18");
            allstu.setSex("男");
            allstu.setPhone("123456789");
            liststu.add(allstu);
        }
        for(Student ss : liststu){
            System.out.println(ss.getId());
        }
        return liststu;
    }

    public String exportStu() {

        List<Student> list = allStu();
        System.out.println("list the size:"+list.size());
        //创建工作簿
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //创建工作表
        XSSFSheet sheet = xssfWorkbook.createSheet();
        xssfWorkbook.setSheetName(0,"学生信息表");
        //创建表头
        XSSFRow head = sheet.createRow(0);
        String[] heads = {"编号","姓名","年龄","性别","手机号"};
        for(int i = 0;i < 5;i++){
            XSSFCell cell = head.createCell(i);
            cell.setCellValue(heads[i]);
        }
        for (int i = 1;i <= list.size();i++) {
            Student student = list.get(i - 1);
            //创建行,从第二行开始，所以for循环的i从1开始取
            XSSFRow row = sheet.createRow(i);
            //创建单元格，并填充数据
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(student.getId());
            cell = row.createCell(1);
            cell.setCellValue(student.getName());
            cell = row.createCell(2);
            cell.setCellValue(student.getAge());
            cell = row.createCell(3);
            cell.setCellValue(student.getSex());
//            cell.setCellValue("男".equals(student.getS_gender().trim())?"男":"女");
            cell = row.createCell(4);
            cell.setCellValue(student.getPhone());
        }
        //创建临时文件的目录
        File file = new File(temp);
        if(!file.exists()){
            file.mkdirs();
        }
        //临时文件路径/文件名
        String downloadPath = file + File.separator  + System.currentTimeMillis() + UUID.randomUUID();
        OutputStream outputStream = null;
        try {
            //使用FileOutputStream将内存中的数据写到本地，生成临时文件
            outputStream = new FileOutputStream(downloadPath);
            xssfWorkbook.write(outputStream);
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if(outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return downloadPath;
    }
}
