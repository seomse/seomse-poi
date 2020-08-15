/*
 * Copyright (C) 2020 Seomse Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
 * ExcelGet 예제
 *
 * @author macle
 */
public class ExcelGetExample {

    private static final Logger logger = LoggerFactory.getLogger(ExcelGetExample.class);

    private ExcelGet excelGet;
    private XSSFRow row;

    /**
     * 엑셀 파일 읽기
     * @param excelFilePath string excel file path
     */
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

    /**
     * cell value string 형태로 얻기
     * @param cellNum int cell num first 0
     * @return string cell value
     */
    private String getCellValue(int cellNum){
        return excelGet.getCellValue(row, cellNum);
    }

    public static void main(String[] args) {
        ExcelGetExample excelGetExample = new ExcelGetExample();
        excelGetExample.load("excel file path");
    }
}
