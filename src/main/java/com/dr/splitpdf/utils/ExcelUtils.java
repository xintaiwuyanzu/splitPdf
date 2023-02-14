package com.dr.splitpdf.utils;

//import com.dr.framework.common.form.core.model.FormData;
//import com.dr.framework.util.DateTimeUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.Assert;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 描述：
 *
 * @author tuzl
 * @date 2020/7/29 18:12
 */
public class ExcelUtils {
    /**
     * 获取第一行中所有标题
     *
     * @param sheet
     * @return
     */
    public static List<String> parse(Sheet sheet) {
        List<String> title = new ArrayList<>();
        Row row1 = sheet.getRow(0);
        Assert.notNull(row1, "无标题内容");
        for (Cell cell : row1) {
            if (!"".equals(cell.getStringCellValue().trim())) {
                title.add(cell.getStringCellValue().trim());
            }
        }
        return title;
    }

    public static LinkedList<Map<String, Object>> readExcel(MultipartFile file, String[] headers) throws Exception {
        String fileName = file.getOriginalFilename();
        boolean isExcel2003 = !fileName.matches("^.+\\.(?i)(xlsx)$");
        InputStream is = null;
        Workbook wb;
        Object value = null;
        Row row;
        Cell cell;

        try {
            is = file.getInputStream();
            if (isExcel2003) {
                wb = new HSSFWorkbook(is);
            } else {
                wb = new XSSFWorkbook(is);
            }

            Sheet sheet = wb.getSheetAt(0);
            LinkedList<Map<String, Object>> linked = new LinkedList();
            for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); ++i) {
                row = sheet.getRow(i);
                if (row != null) {
                    Map<String, Object> map = new HashMap();
                    for (int j = 0; j < headers.length; ++j) {
                        cell = row.getCell(j);
                        if (cell == null) {
                            cell = row.createCell(j);
                        }
                        cell.setCellType(CellType.STRING);
                        map.put(headers[j], null == cell ? "" : cell.getStringCellValue());
                    }
                    linked.add(map);
                }
            }

            return linked;
        } catch (IOException var26) {
            var26.printStackTrace();
            throw new Exception("解析excel异常！");
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException var25) {
                var25.printStackTrace();
            }
        }
    }

    /**
     * @param workbook      工作簿
     * @param sheetName     工作表名
     * @param dataset       正文数据
     * @param headerColumns 标题	数组
     * @param fieldColumns  对象属性字段名数组
     * @Description: 设置工作表（标题与字段名要一一对应）
     */
//    public Sheet creatAuditSheet(SXSSFWorkbook workbook, String sheetName, List<FormData> dataset, List<String> headerColumns, List<String> fieldColumns, String pattern, boolean isExportData)
//            throws IllegalArgumentException {
//        SXSSFSheet sheet = workbook.createSheet(sheetName);
//        sheet.setDefaultColumnWidth(20);
//        SXSSFRow row = sheet.createRow(0);
//        for (int i = 0; i < headerColumns.size(); ++i) {
//            SXSSFCell cellHeader = row.createCell(i);
//            cellHeader.setCellValue(new HSSFRichTextString(headerColumns.get(i)));
//        }
//        if (isExportData) {
//            Iterator<FormData> it = dataset.iterator();
//            int index = 0;
//            Pattern p = Pattern.compile("^//d+(//.//d+)?$");
//            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
//            while (it.hasNext()) {
//                ++index;
//                row = sheet.createRow(index);
//                HashMap<String, Serializable> t = it.next();
//                for (int i = 0; i < fieldColumns.size(); ++i) {
//                    String fieldName = fieldColumns.get(i);
//                    SXSSFCell cell = row.createCell(i);
//                    try {
//                        Object value = t.get(fieldName);
//                        String textValue = null;
//                        if (value instanceof Integer) {
//                            cell.setCellValue((double) (Integer) value);
//                        } else if (value instanceof Float) {
//                            textValue = String.valueOf(value);
//                            cell.setCellValue(textValue);
//                        } else if (value instanceof Double) {
//                            textValue = String.valueOf(value);
//                            cell.setCellValue(textValue);
//                        } else if (value instanceof Long) {
//                            long value2 = (Long) value;
//                            if (value2 > 1000000000000L) {
//                                textValue = DateTimeUtils.longToDate(value2, "yyyy-MM-dd HH:mm:ss");
//                                cell.setCellValue(textValue);
//                            } else {
//                                cell.setCellValue((double) (Long) value);
//                            }
//                        } else if (value instanceof Boolean) {
//                            textValue = "是";
//                            if (!(Boolean) value) {
//                                textValue = "否";
//                            }
//                        } else if (value instanceof Date) {
//                            textValue = sdf.format((Date) value);
//                        } else if (value != null) {
//                            textValue = value.toString();
//                        }
//                        if (textValue != null) {
//                            Matcher matcher = p.matcher(textValue);
//                            if (matcher.matches()) {
//                                cell.setCellValue(Double.parseDouble(textValue));
//                            } else {
//                                HSSFRichTextString richString = new HSSFRichTextString(textValue);
//                                cell.setCellValue(richString);
//                            }
//                        }
//                    } catch (SecurityException var59) {
//                        var59.printStackTrace();
//                    }
//                }
//            }
//        }
//        return sheet;
//    }
}
