package com.ljwm.excel.demo;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

/**
 * JXL 创建Excel
 * Created by yuzhou on 2018/10/20.
 */
@Slf4j
public class JxlExpExcel {

  static final List<String> titles = Arrays.asList("序号", "姓名", "性别");

  /**
   *
   * @param args
   */
  public static void main(String[] args) {
    File file = new File("data/jxl_test.xls");
    try {
      file.getParentFile().mkdirs();
      file.createNewFile();

      // 创建工作簿
      WritableWorkbook workbook = Workbook.createWorkbook(file);

      // 创建sheet
      WritableSheet sheet = workbook.createSheet("Sheet1", 0);

      // 第一行设置列名
      IntStream.range(0, titles.size()).forEach((int i) -> {
        Label label = new Label(i, 0, titles.get(i));
        try {
          sheet.addCell(label);
        } catch (WriteException e) {
          e.printStackTrace();
        }
      });

      // 2到10行追加数据
      IntStream.range(1, 10).forEach((int rowIndex) -> {
        Label label;

        try {
          label = new Label(0, rowIndex, "a" + rowIndex);
          sheet.addCell(label);

          label = new Label(1, rowIndex, "user" + rowIndex);
          sheet.addCell(label);

          label = new Label(2, rowIndex, rowIndex % 2 == 0? "男" : "女");
          sheet.addCell(label);

        } catch (WriteException e) {
          e.printStackTrace();
        }
      });

      workbook.write();
      workbook.close();

    } catch (IOException e) {
      e.printStackTrace();
    } catch (WriteException e) {
      e.printStackTrace();
    }
  }
}
