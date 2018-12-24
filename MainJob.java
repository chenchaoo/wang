package one;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MainJob {

    private Map<String,String> paramMap = new HashMap<>();
    private int param = 0;
    
    public static void main(String[] args) {
        String filePath = ReadConf.getProperties("jobNetFilePath");
        String content = readTxt(filePath);
        String text = replaceBlank(content);
        Map<String, List<Object>> buildListMap = new MainJob().buildListMap(text);
        System.out.println(ExcelRead.generateWorkbook(ReadConf.getProperties("convertFilePath"), "convertFilePath", buildListMap));
    }

    private Map<String,List<Object>> buildListMap(String text) {
        Map<String,List<Object>> map = new HashMap<>();
        List<Object> list = new ArrayList<>();
        String name =text.substring(text.indexOf("="),text.indexOf("{"));
        String unitName = "unit"+name.substring(0,name.indexOf(",,"));
        buildParamMap(name, list);
        int bracketsIndex = 0;
        text = text.substring(text.indexOf("{")+1,text.length()-1);
        Map<String, String> elMap = buildElMap(text, list);
        boolean falg = false;
        int startIndex = 0;
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if('{' == (ch)){
                bracketsIndex += 1;
            }else if('}' == (ch)){
                bracketsIndex -= 1;
                if(bracketsIndex == 0){
                    if(!falg){
                        startIndex = text.indexOf("unit");
                        falg = true;
                    }
                    String unitChild = text.substring(startIndex, i + 1);
                    startIndex = i + 1;
                    String unitChildParam = buildUnitChildParam(elMap, unitChild);
                    StringBuilder sb = insertParamToUnit(unitChild, unitChildParam);
                    Map<String, List<Object>> buildListMap = buildListMap(sb.toString());
                    list.add(buildListMap);
                }
            }
        }
        map.put(unitName, list);
        return map;
    }

    private StringBuilder insertParamToUnit(String unitChild, String unitChildParam) {
        StringBuilder sb = new StringBuilder(unitChild);
        sb.insert(unitChild.indexOf(";"), ","+unitChildParam);
        return sb;
    }

    private String buildUnitChildParam(Map<String, String> elMap, String unitChild) {
        String unitChildName = unitChild.substring(unitChild.indexOf("=")+1,unitChild.indexOf(",,"));
        String unitChildParam = elMap.get(unitChildName);
        return unitChildParam;
    }

    private void buildParamMap(String name, List<Object> list) {
        String propertyName = name.substring(name.indexOf(",,")+2,name.indexOf(";"));
        String[] propertyArr = propertyName.split(",");
        for (int i = 0; i < propertyArr.length; i++) {
            if(paramMap.get(propertyArr[i]) != null){
                list.add(paramMap.get(propertyArr[i])+"="+propertyArr[i]);
            }else{
            	if(i == 0){
            		list.add("ow"+"="+propertyArr[i]);
            		paramMap.put(propertyArr[i], "ow");
            	}else if(i == 1){
            		list.add("gr"+"="+propertyArr[i]);
            		paramMap.put(propertyArr[i], "gr");
            	}else if(propertyArr[i].startsWith("+")){
            		list.add("tmp"+"="+propertyArr[i]);
            		paramMap.put(propertyArr[i], "tmp");
            	}else{
            		String paramStr = String.valueOf(param+=1);
            		list.add("param"+paramStr+"="+propertyArr[i]);
            		paramMap.put(propertyArr[i], "param"+paramStr);
            	}
            }
        }
    }

    private Map<String, String> buildElMap(String text, List<Object> list) {
        String keyValue = "";
        if(text.indexOf("unit") != -1){
            keyValue = text.substring(0,text.indexOf("unit"));
        }else{
            keyValue = text;
        }
        String[] propertyArr2 = keyValue.split(";");
        Map<String,String> elMap = new HashMap<>();
        for (String string : propertyArr2) {
            if(string.startsWith("el=")){
                String key = string.substring(string.indexOf("=")+1,string.indexOf(","));
                String pro = string.substring(string.indexOf(",",string.indexOf(",")+1)+1);
                elMap.put(key, pro);
                list.add(string.substring(0,string.indexOf(",")));
            }else{
                list.add(string);
            }
        }
        return elMap;
    }

    public static String replaceBlank(String str) {
        String dest = "";
        if (str != null) {
            Pattern p = Pattern.compile("\\s*|\t|\r|\n");
            Matcher m = p.matcher(str);
            dest = m.replaceAll("");
        }
        return dest;
    }

    public static String readTxt(String path) {
        StringBuilder content = new StringBuilder("");
        try {
            File file = new File(path);
            InputStream is = new FileInputStream(file);
            InputStreamReader isr = new InputStreamReader(is, "UTF8");
            BufferedReader br = new BufferedReader(isr);
            String str = "";
            while (null != (str = br.readLine())) {
                content.append(str);
            }
            br.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return content.toString();
    }
}
