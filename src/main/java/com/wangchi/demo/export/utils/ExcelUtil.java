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
     * @param keys
     * @param columnNames 列名
     * @return
     */
    public static HSSFWorkbook createWorkBook(List<Map<String, Object>> list, String[] keys, String columnNames[]) {
        // 创建excel工作簿
        HSSFWorkbook wb = new HSSFWorkbook();
        // 创建第一个sheet页，并命名
        HSSFSheet sheet = wb.createSheet(list.get(0).get("sheetName").toString());
        // 设置列宽
        for (int i = 0; i < keys.length; i++) {
            //最后一列为附件URL地址,列宽设置大一些
            if (i == (keys.length - 1)) {
                sheet.setColumnWidth((short) i, (short) (200 * 120));
            } else {
                sheet.setColumnWidth((short) i, (short) (50 * 60));
            }
        }

        // 创建第一行，并设置其单元格格式
        HSSFRow row = sheet.createRow((short) 0);
        row.setHeight((short) 500);
        // 单元格格式(用于列名)
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
        //设置列名
        for (int i = 0; i < columnNames.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(cs);
        }

        //设置首行外,每行每列的值(Row和Cell都从0开始)
        for (short i = 1; i < list.size(); i++) {
            HSSFRow row1 = sheet.createRow((short) i);
            String flag = "";
            //在Row行创建单元格
            for (short j = 0; j < keys.length; j++) {
                HSSFCell cell = row1.createCell(j);
                cell.setCellValue(list.get(i).get(keys[j]) == null ? " " : list.get(i).get(keys[j]).toString());
                if (list.get(i).get(keys[j]) != null) {
                    if ("优".equals(list.get(i).get(keys[j]).toString())) {
                        flag = "优";
                    } else if ("差".equals(list.get(i).get(keys[j]).toString())) {
                        flag = "差";
                    }
                }
            }
            //设置该行样式
            HSSFFont f2 = wb.createFont();
            f2.setFontName("宋体");
            f2.setFontHeightInPoints((short) 10);
            if ("优".equals(flag)) {
                HSSFCellStyle cellStyle = wb.createCellStyle();
                cellStyle.setFont(f2);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
                cellStyle.setLocked(true);
                cellStyle.setWrapText(true);//自动换行
                cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());// 设置背景色
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                //依次为每个单元格设置样式
                for (int m = 0; m < keys.length; m++) {
                    HSSFCell hssfCell = row1.getCell(m);
                    hssfCell.setCellStyle(cellStyle);
                }
            } else if ("差".equals(flag)) {
                HSSFCellStyle cellStyle2 = wb.createCellStyle();
                cellStyle2.setFont(f2);
                cellStyle2.setAlignment(HorizontalAlignment.CENTER);// 左右居中
                cellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
                cellStyle2.setLocked(true);
                cellStyle2.setWrapText(true);//自动换行
                cellStyle2.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());// 设置背景色
                cellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                for (int m = 0; m < keys.length; m++) {
                    HSSFCell hssfCell = row1.getCell(m);
                    hssfCell.setCellStyle(cellStyle2);
                }
            } else {
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
        }
        return wb;
    }

    /**
     *  生成并下载Excel
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
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            ExcelUtil.createWorkBook(list, keys, columnNames).write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        InputStream is = new ByteArrayInputStream(content);
        // 设置response参数
        response.reset();
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + new String((fileName + ".xls").getBytes(), "iso-8859-1"));
        ServletOutputStream out = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
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

