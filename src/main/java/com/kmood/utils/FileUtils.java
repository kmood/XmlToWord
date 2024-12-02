package com.kmood.utils;

import com.google.gson.Gson;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.Base64;
import java.util.HashMap;

public class FileUtils {
	//readToStringByFilepath
	static Gson  gson = JsonBinder.buildNormalBinder("yyyy-MM-dd HH:mm:ss");

	public static String readToStringByFilepath(String filePath)throws IOException {
		StringBuilder data = new StringBuilder();
		Reader reader  =  null;
		try {
			File file = new File(filePath);
			reader = new FileReader(file);
			char [] c = new char [1024];
			int len = 0;
			while ((len = reader.read(c)) != -1){
				data.append(String.copyValueOf(c,0,len));
			}
		} finally {
			if (reader != null)
			reader.close();
		}
		return data.toString();
	}

	public static byte[] readToBytesByFilepath(String filePath)throws IOException {
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		FileInputStream reader  =  null;
		try {
			File file = new File(filePath);
			reader = new FileInputStream(file);
			int len = 0;
			byte[] buffer = new byte[1024];
			while ((len = reader.read(buffer)) != -1){
				output.write(buffer, 0, len);
			}
		} finally {
			if (reader != null)
				reader.close();
		}
		return output.toByteArray();
	}

	public static String getFileSuffixByPath(String path){
		File file = new File(path);
		String fileName = file.getName();
		String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
		return suffix;
	}


	// 判断base64文件的类型
	public static String checkImageBase64Format(String base64ImgData) {
		try{
		byte[] b = Base64.getDecoder().decode(base64ImgData);
		String type = "";
		if (0x424D == ((b[0] & 0xff) << 8 | (b[1] & 0xff))) {
			type = "bmp";
		} else if (0x8950 == ((b[0] & 0xff) << 8 | (b[1] & 0xff))) {
			type = "png";
		} else if (0xFFD8 == ((b[0] & 0xff) << 8 | (b[1] & 0xff))) {
			type = "jpg";
		}else{
			type = "png";
		}
		return type;}catch(Exception ex){
			return "";
		}

	}

	public static void convertBase64DataToImage(String base64ImgData, String filePath) throws IOException {
//		byte[] decode = Base64.getDecoder().decode(base64ImgData);
////		Files.write(Paths.get(filePath), decode, StandardOpenOption.CREATE);
//		FileOutputStream out =new FileOutputStream(filePath);
//		out.write(decode);
//		out.close();

		GenerateImage(base64ImgData,filePath);
	}

	/**
	 * 图片传参占位识别，兼容base64, 如值为 {"type":"pic","url":"路径"} ,将路径转换为
	 * @param param
	 * @return
	 */
	public static String convertParamsToBase64(Object param){
		String paramstr=param!=null?param.toString():"";
		System.out.println(param);
		try{
			HashMap map= gson.fromJson(gson.toJson(param), HashMap.class);
			String type=map.get("type").toString();
			if("pic".equalsIgnoreCase(type)){
				String picurl=map.get("url").toString();
				paramstr=Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(picurl));
			}
		}catch	(Exception e)
		{

		}
		return  paramstr;
	}

	public static String GetImageStr(String imgPath) {// 将图片文件转化为字节数组字符串，并对其进行Base64编码处理
		String imgFile = imgPath;// 待处理的图片
		InputStream in = null;
		byte[] data = null;
		String encode = null; // 返回Base64编码过的字节数组字符串
		// 对字节数组Base64编码
		BASE64Encoder encoder = new BASE64Encoder();
		try {
			// 读取图片字节数组
			in = new FileInputStream(imgFile);
			data = new byte[in.available()];
			in.read(data);
			encode = encoder.encode(data);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				in.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return encode;
	}

	/**
	 * base64字符串转化成图片
	 *
	 * @param imgData
	 *            图片编码
	 * @param imgFilePath
	 *            存放到本地路径
	 * @return
	 * @throws IOException
	 */
	@SuppressWarnings("finally")
	public static boolean GenerateImage(String imgData, String imgFilePath) throws IOException { // 对字节数组字符串进行Base64解码并生成图片
		if (imgData == null) // 图像数据为空
			return false;
		BASE64Decoder decoder = new BASE64Decoder();
		OutputStream out = null;
		try {
			out = new FileOutputStream(imgFilePath);
			// Base64解码
			byte[] b = decoder.decodeBuffer(imgData);
			for (int i = 0; i < b.length; ++i) {
				if (b[i] < 0) {// 调整异常数据
					b[i] += 256;
				}
			}
			out.write(b);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			out.flush();
			out.close();
			return true;
		}
	}



	/**
	 * 递归删除目录下的所有文件及子目录下所有文件
	 * @param dir 将要删除的文件目录
	 * @return boolean Returns "true" if all deletions were successful.
	 *                 If a deletion fails, the method stops attempting to
	 *                 delete and returns "false".
	 */
	public static boolean deleteDir(File dir) {
		if (dir.isDirectory()) {
			String[] children = dir.list();
            //递归删除目录中的子目录下
			for (int i=0; i<children.length; i++) {
				boolean success = deleteDir(new File(dir, children[i]));
				if (!success) {
					return false;
				}
			}
		}
		// 目录此时为空，可以删除
		return dir.delete();
	}


	public static boolean delDir(File file) {
		     if (file.isDirectory()) {
			         File zFiles[] = file.listFiles();
			         for (File file2 : zFiles) {
				            delDir(file2);
				         }
			         file.delete();
			    } else {
			         file.delete();
			    }
			    return true;
		}


}
