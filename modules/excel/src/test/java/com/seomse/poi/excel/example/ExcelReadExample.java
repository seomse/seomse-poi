/*
 * Copyright (C) 2021 Seomse Inc.
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

import com.seomse.poi.excel.ExcelRead;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;

/**
 * @author macle
 */
public class ExcelReadExample extends ExcelRead {
    /**
     * 엑셀 파일 읽기
     * @param excelFilePath string excel file path
     */
    public void load(String excelFilePath){

        try {

            Workbook work = WorkbookFactory.create(new File(excelFilePath));

            excelGet.setWorkbook(work);
            Sheet sheet = work.getSheetAt(0);
            int rowCount = excelGet.getRowCount(sheet);

            for (int i = 0; i < rowCount ; i++) {
                row = sheet.getRow(i);

                int columnCount = excelGet.getColumnCount(row);
                for (int j = 0; j <columnCount ; j++) {
                    System.out.println(getCellValue(j));
                }
            }

        }catch(Exception e){
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        ExcelReadExample excelReadExample = new ExcelReadExample();
        excelReadExample.load("excel file path");
    }
}
