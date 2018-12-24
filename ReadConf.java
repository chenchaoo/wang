package one;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

public class ReadConf {
	
	public static List<String> readTxt(String path) {
        List<String> confList = new ArrayList<>();
        try {
            File file = new File(path);
            InputStream is = new FileInputStream(file);
            InputStreamReader isr = new InputStreamReader(is, "UTF8");
            BufferedReader br = new BufferedReader(isr);
            String str = "";
            while (null != (str = br.readLine())) {
            	confList.add(str);
            }
            br.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return confList;
    }
	
	public static Map<String,Map<String,String>> getConfMap() {
		List<String> confList = readTxt("./conf/conf.txt");
		Map<String,String> confPoList = new HashMap<>();
		Map<String,Map<String,String>> confPoListMap = new HashMap<>();
		String key = "";
		boolean flag = true;
		for (String string : confList) {
			if(string.contains("---------")){
				confPoListMap.put(key, confPoList);
				confPoList = new HashMap<>();
				flag = true;
			}else if(string.indexOf(":") > 0){
				if(flag){
					key = string.substring(0, string.indexOf(":"));
					flag = false;
				}
				confPoList.put(string.substring(0, string.indexOf(":")), string.substring(string.indexOf(":")+1));
			}
		}
		
		if(confPoList.size() != 0){
			confPoListMap.put(key,confPoList);
		}
		return confPoListMap;
	}
	
	public static String getProperties(String key){
		Properties properties = new Properties();
        try {
            InputStream inputStream = new FileInputStream("./conf/filePath.properties");
            BufferedReader bf = new BufferedReader(new InputStreamReader(inputStream, "utf-8"));
            properties.load(bf);
            inputStream.close(); // ¹Ø±ÕÁ÷
        } catch (IOException e) {
            e.printStackTrace();
        }
        return properties.get(key).toString();
	}
}
