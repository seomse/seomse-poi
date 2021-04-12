/*
 * Copyright (C) 2021 Wigo Inc.
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

package com.seomse.poi.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author macle
 */
public class ExcelRead {

    protected ExcelGet excelGet = new ExcelGet();
    protected Sheet sheet;
    protected Row row;


    /**
     * cell value string 형태로 얻기
     * @param cellNum int cell num first 0
     * @return string cell value
     */
    protected String getCellValue(int cellNum){
        return excelGet.getCellValue(row, cellNum);
    }


    /**
     * cell value string 형태로 얻기
     * @param rowNum int row num first 0
     * @param cellNum int cell num first 0
     * @return string cell value
     */
    protected String getCellValue(int rowNum, int cellNum){
        return excelGet.getCellValue(sheet, rowNum, cellNum);
    }


}
