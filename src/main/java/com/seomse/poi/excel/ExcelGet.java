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


package com.seomse.poi.excel;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.SimpleDateFormat;

/**
 * excel 데이터 가져다 쓸때의 유틸성 클래스
 * src test
 * com.seomse.poi.excel.example
 * 위 예제를 보고 활용
 * @author macle
 */
@SuppressWarnings("unused")
public class ExcelGet {

    private FormulaEvaluator formulaEvaluator;

    /**
     * formulaEvaluator 설정
     * @param formulaEvaluator FormulaEvaluator
     */

    public void setFormulaEvaluator(FormulaEvaluator formulaEvaluator){
        this.formulaEvaluator = formulaEvaluator;
    }

    /**
     * XSSFWorkbook 설정
     * 반드시 설정 해야함
     * example)
     * SSFWorkbook work = new XSSFWorkbook(new FileInputStream(excelFilePath));
     * excelGet.setXSSFWorkbook(work);
     * @param xSSFWorkbook XSSFWorkbook
     */
    public void setXSSFWorkbook(XSSFWorkbook xSSFWorkbook){
        formulaEvaluator = xSSFWorkbook.getCreationHelper().createFormulaEvaluator();
    }

    /**
     * cell 값 얻기
     * @param sheet XSSFSheet
     * @param rowNum int row num first 0
     * @param cellNum int cell num first 0
     * @return string
     */
    public String getCellValue( XSSFSheet sheet, int rowNum, int cellNum){
        return getCellValue(sheet, rowNum, cellNum, null);
    }

    /**
     * cell 값 얻기
     * @param row XSSFRow
     * @param cellNum int cell num first 0
     * @return string
     */
    public String getCellValue(XSSFRow row, int cellNum){

        return getCellValue(row, cellNum, null );
    }

    /**
     * cell의 값을 스트링 형태로 반환
     * @param row XSSFRow
     * @param cellNum int cell num first 0
     * @param dateFormat 테이터 포멧(ex:yyyy.MM.dd HH:mm:ss) 날짜형식이 아닐경우 null 전달
     * @return string cell의값을 스트링형태로 반환
     */
    public String getCellValue(XSSFRow row, int cellNum, String dateFormat){
        if(row == null){
            return null;
        }
        return getCellValue(row.getCell(cellNum), dateFormat);
    }

    /**
     * cell의 값을 스트링 형태로 반환
     * @param sheet XSSFSheet
     * @param rowNum int row num first 0
     * @param cellNum int cell num first 0
     * @param dateFormat String java date format example:)yyyyMMdd 날짜 형식이 아닐 경우 null 전달
     * @return string cell의값을 스트링형태로 반환
     */
    public String getCellValue( XSSFSheet sheet, int rowNum, int cellNum, String dateFormat){
        XSSFRow row = sheet.getRow(rowNum);
        if(row == null){
            return null;
        }
        return getCellValue(row.getCell(cellNum), dateFormat);
    }

    /**
     * cell 값 얻기
     * @param cell XSSFCell
     * @return string
     */
    public String getCellValue(XSSFCell cell){
        return getCellValue(cell, null);
    }

    /**
     * cell의 값을 스트링 형태로 반환한다
     * @param cell cell
     * @param dateFormat String java date format example:)yyyyMMdd 날짜 형식이 아닐 경우 null 전달
     * @return string cell의값을 스트링형태로 반환
     */
    @SuppressWarnings("DuplicateBranchesInSwitch")
    public String getCellValue(XSSFCell cell, String dateFormat){
        if(cell == null){
            return null;
        }

        switch(cell.getCellTypeEnum()){


            case NUMERIC:
                return cellNumber(cell, dateFormat);
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue() + "";
            case ERROR:
                return cell.getErrorCellString();
            case BLANK:
                return null;
            case _NONE:
                return null;
            case FORMULA:
                try{
                    switch(formulaEvaluator.evaluateFormulaCellEnum(cell)){
                        case NUMERIC:
                            return cellNumber(cell, dateFormat);
                        case STRING:
                            return cell.getStringCellValue();
                        case BOOLEAN:
                            return cell.getBooleanCellValue() + "";
                        case ERROR:
                            return cell.getErrorCellString();
                        default:
                            return null;
                    }
                }catch(Exception e){
                    switch(cell.getCachedFormulaResultTypeEnum()){
                        case NUMERIC:
                            return cellNumber(cell, dateFormat);
                        case STRING:
                            return cell.getStringCellValue();
                        case BOOLEAN:
                            return cell.getBooleanCellValue() + "";
                        case ERROR:
                            return cell.getErrorCellString();
                        default:
                            return null;
                    }
                }
            default:
                return null;
        }


    }

    /**
     * cell 숫자형 값 얻기
     * @param cell XSSFCell
     * @param dateFormat String java date format example:)yyyyMMdd
     * @return string
     */
    private String cellNumber(XSSFCell cell, String dateFormat){
        if(HSSFDateUtil.isCellDateFormatted(cell) && dateFormat != null){
            SimpleDateFormat formatter = new SimpleDateFormat(dateFormat);
            return formatter.format(cell.getDateCellValue());
        }
        String cellValue = Double.toString(cell.getNumericCellValue());
        if(cellValue.endsWith(".0"))
            cellValue = cellValue.substring(0, cellValue.length()-2);

        return cellValue ;
    }

    
    /**
     * row 개수 얻기
     * poi 사용중 건수가 적게 넘어 와서 개발함
     * @param sheet XSSFSheet
     * @return int row count
     */
    public int getRowCount(XSSFSheet sheet){
        XSSFRow row;
        int rowCount = sheet.getLastRowNum();
        //엑셀 라스트 로우넘 버그처리
        while(true){
            try{
                row = sheet.getRow(rowCount);
                if(row == null){
                    break;
                }
                rowCount ++;
            }catch(Exception e){
                throw new RuntimeException(e);

            }


        }
        return rowCount;
    }

    /**
     * Column 개수 얻기
     * poi 사용중 건수가 적게 넘어 와서 개발함
     * @param row XSSFRow
     * @return int Column count
     */
    public int getColumnCount(XSSFRow row){
        int columnCount = row.getLastCellNum();
        //컬럼 마지막인덱스가져오기 poi자체에대한 버그처리
        while(true){
            try{
                XSSFCell cell = row.getCell(columnCount);

                if(cell == null){
                    break;
                }
                columnCount ++;
            }catch(Exception e){
                throw new RuntimeException(e);
            }
        }
        return columnCount;
    }
}
