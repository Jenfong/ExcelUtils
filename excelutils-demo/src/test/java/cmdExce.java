import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;

import org.apache.log4j.Logger;


public class cmdExce {
	
	
	public static void main(String[] args) {
		Logger.getLogger("cmdExce").info("Start,done!");
		Runtime r = Runtime.getRuntime();
		Process p = null;
		try {
			
			p = r.exec(new String[]{"E:/excelFreezePane.bat","E:/海航酒店集团酒店运营情况综合日报.xls","2","3"});
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} //要执行的外部程序路径及名称
		BufferedReader bf = null;
		try {
			bf = new BufferedReader(new InputStreamReader(p.getInputStream(),"GBK"));
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}     
		String line = " ";
		StringBuffer sb = new StringBuffer();
		try {
			while ((line = bf.readLine()) != null) {
			    sb.append(line);  
			}
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			try {
				bf.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		Logger.getLogger("cmdExce").info("Finished,done!"+sb.toString());
		
	}
	
}
