package com.seomse.poi.excel.example;

import com.seomse.commons.utils.ExceptionUtil;
import com.seomse.poi.excel.ExcelGet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;

/**
 * <pre>
 *  파 일 명 : ExcelGetExample.java
 *  설    명 : ExcelGet 을 활용한 예제 처리
 *
 *  작 성 자 : macle
 *  작 성 일 : 2018.08
 *  버    전 : 1.0
 *  수정이력 :
 *  기타사항 :
 * </pre>
 * @author Copyrights 2018 by ㈜섬세한사람들. All right reserved.
 */
public class ExcelGetExample {

    private static final Logger logger = LoggerFactory.getLogger(ExcelGetExample.class);

    private ExcelGet excelGet;
    private XSSFRow row;

    public void load(String excelFilePath){

        try {
            excelGet = new ExcelGet();
            XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(excelFilePath));
            excelGet.setXSSFWorkbook(work);
            XSSFSheet sheet = work.getSheetAt(0);
            int rowCount = excelGet.getRowCount(sheet);

            for (int i = 0; i < rowCount ; i++) {
                row = sheet.getRow(i);

                int columnCount = excelGet.getColumnCount(row);
                for (int j = 0; j <columnCount ; j++) {
                    System.out.println(getCellValue(j));
                }
            }

        }catch(Exception e){
            logger.error(ExceptionUtil.getStackTrace(e));
        }
    }

    private String getCellValue(int cellNum){
        return excelGet.getCellValue(row, cellNum);
    }

    public static void main(String[] args) {



    }
}
