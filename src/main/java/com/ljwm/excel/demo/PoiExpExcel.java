package com.ljwm.excel.demo;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

/**
 * POI 生成 Excel 文件
 * Created by yuzhou on 2018/10/20.
 */
@Slf4j
public class PoiExpExcel {

  static final List<String> titles = Arrays.asList("序号", "姓名", "性别");

  /**
   * HSSF提供读写Microsoft Excel XLS格式档案的功能。（97-03）
   */
  public static void HSSFExport() {

    // 创建Excel 工作簿
    HSSFWorkbook workbook = new HSSFWorkbook();
    // 创建一个工作表sheet
    HSSFSheet sheet = workbook.createSheet("用户工作表");
    // 创建第一行
    HSSFRow row = sheet.createRow(0);

    IntStream.range(0, titles.size()).forEach((int colIndex) -> {
      HSSFCell cell = row.createCell(colIndex);
      cell.setCellValue(titles.get(colIndex));
    });


    // 2到10行追加数据
    IntStream.range(1, 10).forEach((int rowIndex) -> {
      HSSFRow nextRow = sheet.createRow(rowIndex);
      HSSFCell cell = nextRow.createCell(0);
      cell.setCellValue("a" + rowIndex);

      cell = nextRow.createCell(1);
      cell.setCellValue("user" + rowIndex);

      cell = nextRow.createCell(2);
      cell.setCellValue(rowIndex % 2 == 0? "男" : "女");
    });

    File file = new File("data/poi_hssf_text.xls");
    file.getParentFile().mkdirs();
    try {
      workbook.write(file);
      workbook.close();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  /**
   * XSSF提供读写Microsoft Excel OOXML XLSX格式档案的功能。
   */
  public static void XSSFExport() {

    // 创建Excel 工作簿
    XSSFWorkbook workbook = new XSSFWorkbook();
    // 创建一个工作表sheet
    XSSFSheet sheet = workbook.createSheet("用户工作表");
    // 创建第一行
    XSSFRow row = sheet.createRow(0);


    IntStream.range(0, titles.size()).forEach((int colIndex) -> {
      XSSFCell cell = row.createCell(colIndex);
      cell.setCellValue(titles.get(colIndex));
    });


    // 2到10行追加数据
    IntStream.range(1, 10).forEach((int rowIndex) -> {
      XSSFRow nextRow = sheet.createRow(rowIndex);
      XSSFCell cell = nextRow.createCell(0);
      cell.setCellValue("a" + rowIndex);

      cell = nextRow.createCell(1);
      cell.setCellValue("user" + rowIndex);

      cell = nextRow.createCell(2);
      cell.setCellValue(rowIndex % 2 == 0? "男" : "女");
    });

    File file = new File("data/poi_xssf_text.xlsx");
    file.getParentFile().mkdirs();
    try {
      file.createNewFile();
      OutputStream outputStream = FileUtil.getOutputStream(file);
      workbook.write(outputStream);
      outputStream.close();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static void XSSFExport2() {

    //可以表示xls和xlsx格式文件的类
    XSSFWorkbook  workbook = new XSSFWorkbook();
    try {
      //新创建的xls需要新创建新的工作簿，offine默认创建的时候会默认生成三个sheet
      Sheet sheet = workbook.createSheet("first sheet");
      FileOutputStream out = new FileOutputStream("data/createWorkBook.xlsx");
      workbook.write(out);
      out.close();
      System.out.println("createWorkBook success");
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static void main(String[] args) {

    HSSFExport();
    log.info("HSSF Export Done!!!");
    XSSFExport();
    XSSFExport2();
    log.info("XSSF Export Done!!!");
  }
}
