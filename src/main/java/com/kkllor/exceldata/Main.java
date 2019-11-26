package com.kkllor.exceldata;

import com.kkllor.exceldata.entity.ResultBean;

import java.io.IOException;
import java.util.List;

public class Main {
    public static String[] SYMBOL1 = {"<", "=", ">"};

    public static String[] SYMBOL2 = {"=<", "<=", "=>", ">="};


    public static String FILENAME = "原始数据.xlsx";

    public static void main(String[] args) throws IOException {
//if 3.3=< C20_4=<4.5 then   C20_4group=2;
        List<ResultBean> resultBeanList = ReadWriteExcelFile.readXLSXFile(FILENAME);
        for (ResultBean resultBean : resultBeanList) {
            analyse(resultBean);
            System.out.println(resultBean.toString());
        }

        ReadWriteExcelFile.writeXLSXFile("结果.xlsx", resultBeanList);
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

//        List<String> symbolList = new ArrayList<>();

        int symbolCount = 0;

        for (String symbol : SYMBOL2) {
            symbolCount += indexSubString(symbol, leftStr);

            if (symbolCount > 0) {
//                symbolList.add(symbol);
                leftStr = leftStr.replaceAll(symbol, "");
                tmpLeftStr = tmpLeftStr.replaceAll(symbol, "^");
            }
        }

        for (String symbol : SYMBOL1) {

            int count = indexSubString(symbol, leftStr);
            symbolCount += count;
            if (count > 0) {
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
                resultBean.setRange(split3[1]);
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
