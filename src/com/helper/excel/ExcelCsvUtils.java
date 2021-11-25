package com.helper.excel;

import com.opencsv.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelCsvUtils {

    public static List<String[]> getCsvData(InputStream in, String charsetName) {
        List<String[]> list = new ArrayList<>();
        int i = 0;
        String[] title;
        try (CSVReader csvReader = new CSVReaderBuilder(new BufferedReader(new InputStreamReader(in, charsetName))).build()) {
            List<String[]> writeList = new ArrayList<>();
            int index = -1;
            for (String[] next : csvReader) {
                if (i++ == 0) {
                    title = next;
                    index = getParamIndex(title);
                    writeList.add(title);
                    continue;
                }
                if (index == -1) throw new IllegalArgumentException("找不到目标列");
                List<String> params = plug(next, index);
                addToList(next, index, params, writeList);
            }
            return writeList;
        } catch (Exception e) {
            System.out.println("CSV文件读取异常");
            e.printStackTrace();
            return list;
        }
    }

    private static List<String> plug(String[] next, int index) {
        String string = next[index];
        List<String> list = new ArrayList<>();
        if (string.charAt(0) == '[') {
            string = string.substring(1, string.length() - 1);
        } else {
            list.add(string);
            return list;
        }
        int cut = 0;
        while (cut < string.length() && cut != -1) {
            int cutNext = string.indexOf("},{", cut);
            if (cutNext != -1) {
                list.add(string.substring(cut, cutNext + 1));
                cut = cutNext + 2;
            } else {
                list.add(string.substring(cut));
                break;
            }
        }
        return list;
    }

    private static void addToList(String[] next, int index, List<String> params, List<String[]> writeList) {
        for (String param : params) {
            String[] write = next.clone();
            write[index] = param;
            writeList.add(write);
        }
    }

    private static int getParamIndex(String[] title) {
        for (int i = 0; i < title.length; i++) {
            if (title[i].equals("eee")) return i;
        }
        return -1;
    }

    public static void parseCsv(String filePath) {
        parseCsv(filePath, "gbk");
    }

    public static void parseCsv(String filePath, String charset) {
        try (FileInputStream in = new FileInputStream(filePath)) {
            List<String[]> csvData = getCsvData(in, charset);
            writeToCsv(filePath, charset, csvData);
        } catch (IOException e) {
            // no-op
            e.printStackTrace();
        }
    }

    private static void writeToCsv(String filePath, String charsetName, List<String[]> csvData) {
        int index = filePath.indexOf(".csv");
        String path = filePath.substring(0, index) + "_write" + filePath.substring(index);
        File file = new File(path);
        if (file.exists()) {
            boolean delete = file.delete();
            System.out.println("delete " + delete);
        }
        try (ICSVWriter csvWriter = new CSVWriterBuilder(
            new BufferedWriter(new OutputStreamWriter(new FileOutputStream(path), charsetName))
        ).build()) {
            csvWriter.writeAll(csvData);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
