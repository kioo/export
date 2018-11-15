package com.wangchi.demo.export.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;
import java.util.Map;

/**
 * Excel 导出工具类
 */
public class ExcelUtil {

    /**
     * 创建excel文档
     *
     * @param list
     * @param keys        列名称集合
     * @param columnNames 列名
     * @return
     */
    public static HSSFWorkbook createWorkBook(List<Map<String, Object>> list, String[] keys, String columnNames[]) {
        // 1. 创建excel工作簿
        HSSFWorkbook wb = new HSSFWorkbook();
        // 2. 创建第一个sheet页，并命名
        HSSFSheet sheet = wb.createSheet(list.get(0).get("sheetName").toString());
        // 3. 设置每列的宽
        for (int i = 0; i < keys.length; i++) {
            sheet.setColumnWidth((short) i, (short) (50 * 60));
        }
        // 4. 创建第一行，设置其单元格格式，并将数据放入
        HSSFRow row = sheet.createRow((short) 0);
        row.setHeight((short) 500);
        // 4.1 设置单元格格式
        HSSFCellStyle cs = wb.createCellStyle();
        HSSFFont f = wb.createFont();
        f.setFontName("宋体");
        f.setFontHeightInPoints((short) 10);
        f.setBold(true);
        cs.setFont(f);
        cs.setAlignment(HorizontalAlignment.CENTER);// 水平居中
        cs.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中
        cs.setLocked(true);
        cs.setWrapText(true);//自动换行
        // 4.2 设置列名（取出列名集合进行创建）
        for (int i = 0; i < columnNames.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(cs);
        }

        // 5. 设置首行外,每行每列的值(Row和Cell都从0开始)
        for (short i = 1; i < list.size(); i++) {
            HSSFRow row1 = sheet.createRow((short) i);
            String flag = "";
            // 5.1 在Row行创建单元格
            for (short j = 0; j < keys.length; j++) {
                HSSFCell cell = row1.createCell(j);
                cell.setCellValue(list.get(i).get(keys[j]) == null ? " " : list.get(i).get(keys[j]).toString());
            }
            // 5.2 设置该行样式
            HSSFFont f2 = wb.createFont();
            f2.setFontName("宋体");
            f2.setFontHeightInPoints((short) 10);
            // 5.3 设置单元格样式
            HSSFCellStyle cs2 = wb.createCellStyle();
            cs2.setFont(f2);
            cs2.setAlignment(HorizontalAlignment.CENTER);// 左右居中
            cs2.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
            cs2.setLocked(true);
            cs2.setWrapText(true);//自动换行
            for (int m = 0; m < keys.length; m++) {
                HSSFCell hssfCell = row1.getCell(m);
                hssfCell.setCellStyle(cs2);
            }

        }
        return wb;
    }

    /**
     * 生成并下载Excel
     *
     * @param list
     * @param keys
     * @param columnNames
     * @param fileName
     * @param response
     * @throws IOException
     */
    public static void downloadWorkBook(List<Map<String, Object>> list,
                                        String keys[],
                                        String columnNames[],
                                        String fileName,
                                        HttpServletResponse response) throws IOException {
        // 1. 声明字节输出流
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            // 2. 生成excel文件并写入输出流
            ExcelUtil.createWorkBook(list, keys, columnNames).write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 3. 将输出流转换成byte[] 数组
        byte[] content = os.toByteArray();
        // 4. 将数组放入输入流中
        InputStream is = new ByteArrayInputStream(content);
        // 5. 设置response参数
        response.reset(); // 重置response的设置
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + new String((fileName + ".xls").getBytes(), "iso-8859-1"));
        // 6. 创建Servlet 输出流对象
        ServletOutputStream out = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            // 6.1装载缓冲输出流
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
            // 6.2 输出内容
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (final IOException e) {
            throw e;
        } finally {
            if (bis != null)
                bis.close();
            if (bos != null)
                bos.close();
        }
    }
}

