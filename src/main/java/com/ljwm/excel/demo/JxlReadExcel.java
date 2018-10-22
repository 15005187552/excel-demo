package com.ljwm.excel.demo;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.util.stream.IntStream;

/**
 * JXL 读取Excel 数据
 * Created by yuzhou on 2018/10/20.
 */
public class JxlReadExcel {

  public static void main(String[] args) {

    try {
      Workbook workbook = Workbook.getWorkbook(new File("data/jxl_test.xls"));

      Sheet sheet = workbook.getSheet(0);

      int columns = sheet.getColumns();
      int rows = sheet.getRows();

      IntStream.range(0, rows).forEach((int rowIndex) -> {
        IntStream.range(0, columns).forEach((int colIndex) -> {
          Cell cell = sheet.getCell(colIndex, rowIndex);
          System.out.print(cell.getContents() + "\t");
        });
        System.out.println();
      });


      workbook.close();
    } catch (IOException e) {
      e.printStackTrace();
    } catch (BiffException e) {
      e.printStackTrace();
    }
  }
}
