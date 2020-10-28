package org.poi;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

/**
 * @author y
 * @create 2018-01-19 0:13
 * @desc
 **/
public class ExcelReaderUtil {
	//excel2003扩展名
	public static final String EXCEL03_EXTENSION = ".xls";
	//excel2007扩展名
	public static final String EXCEL07_EXTENSION = ".xlsx";

	/**
	 * 遍历出来的数据信息
	 */
	List<List<String>> list = new ArrayList<List<String>>();
	/**
	 * 每获取一条记录，即打印
	 * 在flume里每获取一条记录即发送，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
	 * @param sheetName
	 * @param sheetIndex
	 * @param curRow
	 * @param cellList
	 */
	public static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
			StringBuffer oneLineSb = new StringBuffer();
			oneLineSb.append(filePath);
			oneLineSb.append("--");
			oneLineSb.append("sheet" + sheetIndex);
			oneLineSb.append("::" + sheetName);//加上sheet名
			oneLineSb.append("--");
			oneLineSb.append("row" + curRow);
			oneLineSb.append("::");
			for (String cell : cellList) {
				oneLineSb.append(cell.trim());
				oneLineSb.append("|");
			}
			String oneLine = oneLineSb.toString();
			if (oneLine.endsWith("|")) {
				oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
			}// 去除最后一个分隔符
		    System.out.println("hehe");
			System.out.println(oneLine);
	}

    public static List<List<String>> readExcel(String fileName) throws Exception {

        Boolean Excel = true;
        ExcelXlsxReaderWithDefaultHandler excelXlsxReader = null;
        ExcelXlsReader excelXls = null;

        InputStream is = new FileInputStream(fileName);
        if (!is.markSupported()) {
            is = new PushbackInputStream(is, 8);
        }

        if (POIFSFileSystem.hasPOIFSHeader(is)) { //处理excel2003文件
            Excel = true;
            excelXls = new ExcelXlsReader();
            excelXls.process(fileName);

        } else if (POIXMLDocument.hasOOXMLHeader(is)) {//处理excel2007文件
            Excel = false;
            excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            excelXlsxReader.process(fileName);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
        List<List<String>> data = null;
        if (Excel) {//2003放回的数据
            data = excelXls.Data();
        } else {//2007返回的数据
            data = excelXlsxReader.Data();
        }
        is.close();
        return data;

    }

	public static void main(String[] args) throws Exception {
//		String path="http://192.168.8.252/group1/M00/00/58/oYYBAF-JZ86EVIB1AAAAAMz2I9o951.xls";
		String path="C:\\Users\\nl\\Desktop\\11.xls";
//		BufferedReader br = null;
/*		HttpURLConnection httpUrl = null;
		URL url = new URL(path);
		httpUrl = (HttpURLConnection) url.openConnection();
		httpUrl.connect();
		File file = getFileByUrl(path);
		String absolutePath = file.getAbsolutePath();*/
		ExcelReaderUtil.readExcel(path);

	}

	//url转file
	private static File getFileByUrl(String fileUrl) {
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		BufferedOutputStream stream = null;
		InputStream inputStream = null;
		File file = null;
		try {
			URL imageUrl = new URL(fileUrl);
			HttpURLConnection conn = (HttpURLConnection) imageUrl.openConnection();
			conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");
			conn.connect();
			inputStream = conn.getInputStream();
			byte[] buffer = new byte[1024];
			int len = 0;
			while ((len = inputStream.read(buffer)) != -1) {
				outStream.write(buffer, 0, len);
			}
			file = File.createTempFile("file", fileUrl.substring(fileUrl.lastIndexOf("."), fileUrl.length()));
			FileOutputStream fileOutputStream = new FileOutputStream(file);
			stream = new BufferedOutputStream(fileOutputStream);
			stream.write(outStream.toByteArray());
		} catch (Exception e) {
		} finally {
			try {
				if (inputStream != null) {
					inputStream.close();
				}
				if (stream != null) {
					stream.close();
				}
				outStream.close();
			} catch (Exception e) {
			}
		}
		return file;
	}

}
