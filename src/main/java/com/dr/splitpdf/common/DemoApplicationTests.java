package com.dr.splitpdf.common;

/**
 * @Author: caor
 * @Date: 2023-02-13 19:35
 * @Description:
 */

import com.dr.splitpdf.utils.PdfUtil;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class DemoApplicationTests {
    @Test
    void contextLoads() throws Exception {
//        PdfUtil.splitPdf("C:\\Users\\bluel\\Desktop\\0003-1912-D10-GL-0001.pdf", 1);
        PdfUtil.splitPdf("D:\\ceshi", "D:\\ceshi\\pdf","");
    }
}