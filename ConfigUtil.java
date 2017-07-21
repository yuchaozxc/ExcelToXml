package com.yuchao.cn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class ConfigUtil {
	//获取一个文本输入流，该文本输入流所输入的是config.properties配置文件
	public FileInputStream getFileInputStream(){
		//声明一个文本输入流
		FileInputStream in = null;
		try {
			//实例化文本输入流，输入config.properties配置文件
			in = new FileInputStream("conf/config.properties");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//返回该文本输入流
		return in;
	}
	//关闭一个文本输入流
	public void closeFileInputStream(FileInputStream in){
		try {
			//关闭文本输入流
			in.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	//通过Properties获取名称为pathName的值
	public String getPath(Properties pro, String pathName) {
		//获取本项目的根目录
		String firstPath = ExcelToXml.class.getProtectionDomain().getCodeSource().getLocation().getPath();
		//获取本项目根目录的上一级目录
		firstPath = firstPath.substring(0, firstPath.lastIndexOf("/"));
		//获取config.properties中配置的名称为pathName的属性值
		String lastPath = pro.getProperty(pathName);
		//判断lastPath是否是以..开头的
		if (lastPath.substring(0, 2).equals("..")) {
			//如果lastPath是以..开头的，我们就调用该方法获取它的绝对路径
			lastPath = getPath(firstPath, lastPath);
		} else {
			//如果不是以..开头的，我们再去判断是否是以.开头的
			if (lastPath.substring(0, 1).equals(".")) {
				//如果是以.开头的，我们在这里处理，获取他的绝对路径
				lastPath = firstPath + lastPath.substring(1, lastPath.length());
			}
		}
		//返回该绝对路径
		return lastPath;
	}
	//当配置中的路径是以..开头的时候，我们通过该方法可以获取到我们要的文件夹的绝对路径
	public String getPath(String firstPath, String lastPath) {
		//判断lastPath中是否存在..
		if (lastPath.indexOf("..") != -1) {
			//将firstPath进行截取，去掉firstPath的最后一个目录结构
			firstPath = firstPath.substring(0, firstPath.lastIndexOf("/"));
			//将lastPath进行截取，去掉lastPath的前3个字符，（../）
			lastPath = lastPath.substring(3, lastPath.length());
			//递归调用
			return getPath(firstPath, lastPath);
		} else {
			//当lastPath中不存在..的时候，我们将firstPath和lastPath进行拼接，获得绝对路径
			return firstPath + "/" + lastPath;
		}
	}
}
