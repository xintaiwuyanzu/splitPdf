package com.dr.splitpdf.utils;

/**
 * @Author: caor
 * @Date: 2023-02-13 19:34
 * @Description:
 */

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import lombok.extern.slf4j.Slf4j;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * PDF处理工具类
 *
 * @author azhuzhu 2020/8/11 14:53
 */
@Slf4j
public class PdfUtil {

    /**
     * 将PDF文件切分成多个PDF
     *
     * @param filename  文件名
     * @param splitSize 拆分单个文件页数
     */
    public static void splitPdf(String filename, int splitSize) throws Exception {
        PdfReader reader;
        try {
            reader = new PdfReader(filename);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new Exception("读取PDF文件失败");
        }
        int numberOfPages = reader.getNumberOfPages();
        int newFileCount = 0;
        // PageNumber是从1开始计数的
        int pageNumber = 1;
        while (pageNumber <= numberOfPages) {
            Document doc = new Document();
            String splitFileName = filename.substring(0, filename.length() - 4) + "(" + newFileCount + ").pdf";
            PdfCopy pdfCopy;
            try {
                pdfCopy = new PdfCopy(doc, new FileOutputStream(splitFileName));
            } catch (FileNotFoundException | DocumentException e) {
                log.error(e.getMessage(), e);
                throw new Exception("切割文件副本创建失败");
            }
            doc.open();
            // 将pdf按页复制到新建的PDF中
            for (int i = 1; pageNumber <= numberOfPages && i <= splitSize; ++i, pageNumber++) {
                doc.newPage();
                PdfImportedPage page = pdfCopy.getImportedPage(reader, pageNumber);
                pdfCopy.addPage(page);
            }
            doc.close();
            newFileCount++;
            pdfCopy.close();
        }
    }


    /**
     * @param excelDir  excel目录文件夹
     * @param pdfDir    pdf文件夹
     * @param newPdfDir 拆分后端newPdf文件夹（可以为空，为空的话把文件按照档号创建一个文件夹，把拆分后端文件放到该文件夹中，文件夹跟pdfDir文件夹同级别）
     * @throws Exception
     */
    public static void splitPdf(String excelDir, String pdfDir, String newPdfDir) throws Exception {
        //读取excelDir中excel列表（xls,xlsx）
        //读取excel列表中数据 excel列表中第一行必须包含档号、页数、页号，这几个顺序无所谓
        //读取pdfDir列表中pdf文件
        //根据excel中档号like查找pdf文件
        //根据档号对应的页数和页号进行拆分pdf,并输出到newPdfDir中
        //在每一个excel最后一列补充拆分状态：成功，或者失败原因（例如：未找到pdf等）
        throw new Exception("读取PDF文件失败");
    }

}

