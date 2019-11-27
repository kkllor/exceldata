package com.kkllor.exceldata;

import com.kkllor.exceldata.entity.ResultBean;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

public class Main {
    public static String[] SYMBOL1 = {"<", "=", ">"};

    public static String[] SYMBOL2 = {"=<", "<=", "=>", ">="};
//    public static String FILENAME = "原始数据.xlsx";

    public static ExecutorService pool = Executors.newFixedThreadPool(Runtime.getRuntime().availableProcessors() * 2);

    public static void main(String[] args) throws IOException {
        File dir = new File(".");
        File[] files = dir.listFiles();

        for (File file : files) {
            String fileName = file.getName();
            pool.execute(() -> {
                if (fileName.contains(".xlsx")) {
                    process(fileName);
                }
            });
        }
        pool.shutdown();
    }

    public static void process(String fileName) {
        List<ResultBean> resultBeanList = null;
        try {
            resultBeanList = ReadWriteExcelFile.readXLSXFile(fileName);
            for (ResultBean resultBean : resultBeanList) {
                long start = System.currentTimeMillis();
                analyse(resultBean);
                System.out.println("analyse lasts:" + (System.currentTimeMillis() - start));
                System.out.println(resultBean.toString());
            }

            File dir = new File("results");
            if (!dir.exists()) {
                dir.mkdirs();
            }
            ReadWriteExcelFile.writeXLSXFile("results/" + fileName, resultBeanList);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void analyse(ResultBean resultBean) {
        String originData = resultBean.getOriginData();
        if (originData == null) {
            return;
        }

        if (!(originData.contains("if") && originData.contains("then"))) {
            return;
        }
        originData = originData.replaceFirst("if", "");
        String[] split1 = originData.split("then");
        if (split1.length != 2) {
            return;
        }

        String[] split2 = split1[1].split("=");
        if (split2.length != 2) {
            return;
        }
        resultBean.setName(split2[0]);
        resultBean.setType(split2[1].replaceAll(";", ""));

        String leftStr = split1[0];
        String tmpLeftStr = leftStr;

        List<String> symbolList = new ArrayList<>();

        int symbolCount = 0;

        for (String symbol : SYMBOL2) {
            symbolCount += indexSubString(symbol, leftStr);

            if (symbolCount > 0) {
                symbolList.add(symbol);
                leftStr = leftStr.replaceAll(symbol, "");
                tmpLeftStr = tmpLeftStr.replaceAll(symbol, "^");
            }
        }

        for (String symbol : SYMBOL1) {

            int count = indexSubString(symbol, leftStr);
            symbolCount += count;
            if (count > 0) {
                symbolList.add(symbol);
                leftStr = leftStr.replaceAll(symbol, "");
                tmpLeftStr = tmpLeftStr.replaceAll(symbol, "^");
            }
        }
        if (symbolCount == 2) {
            String[] split3 = tmpLeftStr.split("\\^");
            if (split3.length == 3) {
                resultBean.setRange(split3[0] + "~" + split3[2]);
            }
        } else if (symbolCount == 1) {
            String[] split3 = tmpLeftStr.split("\\^");
            if (split3.length == 2) {
                resultBean.setRange(symbolList.get(0) + split3[1]);
            }
        }
    }


    public static int indexSubString(String subString, String data) {
        int count = 0;
        int length = data.length();
        for (int i = 0; i < length; ) {
            i = data.indexOf(subString, i);
            if (i == -1) {
                break;
            }
            count++;
            i++;
        }
        return count;

    }

}
