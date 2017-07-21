package com.yuchao.cn;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToXml {
	public static void main(String[] args) throws Exception {
		ExcelToXml excelToXml = new ExcelToXml();
		excelToXml.excelToXml();
	}
	
	@SuppressWarnings("unchecked")
	public void excelToXml(){
		//实例化ConfigUtil类，用来获取config.properties中配置的属性
		ConfigUtil configUtil = new ConfigUtil();
		//打开一个文件输入流，这个输入流读取了config.properties
		FileInputStream in = configUtil.getFileInputStream();
		//实例化一个properties配置文件的工具类
		Properties pro = new Properties();
		try {
			//载入文件输入流中的数据
			pro.load(in);
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		//获得ExcelPath的路径，该路径在config.properties中可以配置成相对路径，也可以配置成绝对路径，
		//但是，如果是相对路径，这个路径是相对于本项目根目录的，不是相对于本类文件的
		String excelPath = configUtil.getPath(pro,"ExcelPath");
		//获得ClientPath的路径，该路径在config.properties中可以配置成相对路径，也可以配置成绝对路径，
		//但是，如果是相对路径，这个路径是相对于本项目根目录的，不是相对于本类文件的
		String clientPath = configUtil.getPath(pro,"ClientPath");
		//获得ClientPath的路径，该路径在config.properties中可以配置成相对路径，也可以配置成绝对路径，
		//但是，如果是相对路径，这个路径是相对于本项目根目录的，不是相对于本类文件的
		String serverPath = configUtil.getPath(pro,"ServerPath");
		//获得需要在生成的xml文件中声明的版本号，字符集的配置
		String xmlDeclaration= pro.getProperty("xmlDeclaration");
		System.out.println("ExcelPath = " + excelPath);
		System.out.println("ClientPath = " + clientPath);
		System.out.println("ServerPath = " + serverPath);
		System.out.println("XmlDeclaration = "+xmlDeclaration);
		//文件输入流使用完毕，关闭这个文件输入流
		configUtil.closeFileInputStream(in);
		//读取Excel的存放路径，这个路径需要指向一个已存在的文件夹，不然会报错，因为要根据这里面的excel来生成xml，所以，如果文件夹不存在，我不会主动去创建
		File file = new File(excelPath);
		//判断该file是不是一个文件夹
		if (file.isDirectory()) {
			//获取文件夹中的所有文件
			File[] excelFiles = file.listFiles();
			//遍历这些文件
			for (File excelFile : excelFiles) {
				//判断是否是文件
				if (excelFile.isFile()) {
					//获取文件名（包括后缀名）
					String excelName = excelFile.getName();
					//获取文件名（去掉后缀名）
					String firstName = excelName.substring(0, excelName.lastIndexOf("."));
					//获取文件的后缀名
					String excelLastName = excelName.substring(excelName.lastIndexOf(".") + 1, excelName.length());
					//声明一个workbook，用来操作我们要操作的excel文档用的
					Workbook wb = null;
					//声明一个文本输入流，用来读取我们要操作的excel文档用的
					FileInputStream fis = null;
					try {
						//实例化文本输入流
						fis = new FileInputStream(excelFile);
					} catch (FileNotFoundException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
					//声明一个sheet,用来操作我们要操作的sheet
					Sheet sheet = null;
					//声明一个row，这个row用来操作sheet中的第一列，这里的规则第一列存放的是字段名
					Row namesRow = null;
					//声明一个row，这个row用来操作sheet中的第二列，这里的规则第二列存放的是字段说明
					Row annotationRow = null;
					//声明一个row,这个row用来操作sheet中的第三列，这里的规则第三列存放的是字段的类型
					Row classRow = null;
					//声明一个row,这个row用来操作sheet中的第四列，这里的规则第四列存放的是一个标志位，
					//标志这个字段生成的时候是放入client使用的xml文件中的，还是放入server使用的xml文件中的，
					//C表示放入client使用的xml文件中，S表示放入server使用的xml文件中，A表示在client和server使用的xml文件中都放入
					Row typeRow = null;
					//我们每一张excel表中必须要有一个叫Key的字段，这个字段是确定数据的唯一性的，相当于id，这里声明的keyNum是表示Key这个字段在sheet中的第几列
					int keyNum = 0;
					//声明一个boolean值，该值表示是否要生成client端使用的xml文件，当key的标志位为C的时候只生成client端使用的xml文件，当key的标志位为S的时候，
					//只生成server端使用的xml文件，当key的标志位为A的时候，表示既要生成client端使用的xml文件，又要生成server端使用的xml文件
					boolean cfal = false;
					//声明一个boolean值，该值表示是否要生成server端使用的xml文件
					boolean sfal = false;
					//声明一个List，用来存放所有的字段名
					List<String> nameList = null;
					//声明一个List，用来存放所有的字段说明
					List<String> annotationList = null;
					//声明一个List，用来存放所有的字段类型
					List<String> classList = null;
					//声明一个List，用来存放所有的要放入client端使用的xml文件的字段的位置
					List<Integer> cnums = null;
					//声明一个List，用来存放所有的要放入server端使用的xml文件的字段的位置
					List<Integer> snums = null;
					//实例化一个xml的名称，所生成的xml文件就叫这个名字
					String xmlName = firstName + ".xml";
					//声明一个xml文件，这个文件就是我们要生成的xml文件
					File xmlFile = null;
					//声明一个字符串，这个字符传存放的是我们要放入xml文件中的内容
					String outputStr = "";
					//判断该文件的后缀名是否是xls结尾的，主要是为了区分excel的版本
					if (excelLastName.equals("xls")) {
						POIFSFileSystem fs = null;
						try {
							fs = new POIFSFileSystem(fis);
							//实例化workbook
							wb = new HSSFWorkbook(fs);
						} catch (IOException e) {
							e.printStackTrace();
						}
					//判断该文件的后缀名是否是xlsx结尾的，主要是为了区分excel的版本
					} else if(excelLastName.equals("xlsx")){
						try {
							//实例化workbook
							wb = new XSSFWorkbook(fis);
						} catch (IOException e) {
							e.printStackTrace();
						}
					//不是excle文件就跳过本次循环
					}else{
						continue;
					}
					//实例化sheet，这里我默认取的是文件中的第一个sheet，大家也可以改成用sheet名来取的，wb.getSheet("sheet名");
					sheet = wb.getSheetAt(0);
					//获取sheet中的第一行，也就是字段名那一行
					namesRow = sheet.getRow(0);
					//获取第一行的内容
					Object[] obj = getNames(namesRow);
					//将第一行的内容赋值给nameList
					nameList = (List<String>) (obj[0]);
					//获得key在excel表中的哪一列
					keyNum = (int) (obj[1]);
					//判断，如果第一行为空，就跳过本次循环
					if (nameList == null || nameList.size() == 0) {
						continue;
					}
					//获得sheet中的第二行，也就是字段说明那一行
					annotationRow = sheet.getRow(1);
					//获得字段说明的内容
					annotationList = getAnnotations(annotationRow);
					//判断，如果第二行为空，就跳过本次循环
					if (annotationList == null || annotationList.size() == 0) {
						continue;
					}
					//获得sheet中的第三行，也就是字段类型那一行
					classRow = sheet.getRow(2);
					//获得字段类型的内容
					classList = getClasses(classRow);
					//判断，如果第三行为空，就跳过本次循环
					if (classList == null || classList.size() == 0) {
						continue;
					}
					//获得sheet中的第四行，也就是标志位那一行
					typeRow = sheet.getRow(3);
					//获得标志位的信息
					Object[] tobj = getTypes(typeRow, keyNum);
					//获得哪些列是要放入到client端使用的xml文件中的
					cnums = (List<Integer>) tobj[0];
					//获得哪些列是要放入到server端使用的xml文件中的
					snums = (List<Integer>) tobj[1];
					//获取是否生成客户端xml文件
					cfal = (boolean) tobj[2];
					//获取是否生成server端使用的xml文件
					sfal = (boolean) tobj[3];
					//判断是否生成client端使用的xml文件
					if (cfal) {
						//获取要向xml文件中打印的内容
						outputStr = getOutputStr(nameList, annotationList, classList, firstName, sheet, cnums,xmlDeclaration, false);
						System.out.println(outputStr);
						//实例化client端使用的xml文件
						xmlFile = new File(clientPath + "/" + xmlName);
						try {
							//将内容写入到client端使用的xml文件中
							writer(xmlFile, outputStr);
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}
					//判断是否生成server端使用的xml文件
					if (sfal) {
						//获取要向xml文件中打印的内容
						outputStr = getOutputStr(nameList, annotationList, classList, firstName, sheet, snums,xmlDeclaration, true);
						System.out.println(outputStr);
						//实例化server端使用的xml文件
						xmlFile = new File(serverPath + "/" + xmlName);
						try {
							//将内容写入到server端使用的xml文件中
							writer(xmlFile, outputStr);
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}

				}
			}
		}
	}
	//获取excel中第一行所包含的信息
	private Object[] getNames(Row namesRow) {
		//实例化一个List,该list中存放的是所有的字段名
		List<String> nameList = new ArrayList<String>();
		//实例化一个int值，该值表示key字段在excel中的位置
		int keyNum = 0;
		//实例化一个object类型的数组，该数组中存放的是所有的字段名和key字段的位置
		Object[] obj = new Object[2];
		//判断namesRow这个行是否为空
		if (namesRow != null) {
			//遍历namesRow这一行
			for (int i = 0; i < namesRow.getLastCellNum(); i++) {
				//获取单元格
				Cell cell = namesRow.getCell(i);
				//判断单元格是否为空
				if (cell != null) {
					//添加单元格的内容到nameList中
					nameList.add(cell.getStringCellValue());
					//判断这个单元格的内容是不是Key
					if (cell.getStringCellValue().equalsIgnoreCase("Key")) {
						//记录Key的位置
						keyNum = i;
					}
				}
			}
		}
		//将所有的字段名放入obj[0]
		obj[0] = nameList;
		//将key列的位置放入obj[1]
		obj[1] = keyNum;
		//返回obj
		return obj;
	}
	//获取字段说明那一行的数据
	private List<String> getAnnotations(Row annotationRow) {
		//声明一个list，用来存放所有的字段说明
		List<String> annotationList = new ArrayList<String>();
		//判断，字段说明那一行是否为空
		if (annotationRow != null) {
			//遍历字段说明这一行所有的单元格
			for (int i = 0; i < annotationRow.getLastCellNum(); i++) {
				//获取单元格
				Cell cell = annotationRow.getCell(i);
				//判断单元格是否为空
				if (cell != null) {
					//将单元格中的内容放入List中
					annotationList.add(cell.getStringCellValue());
				}
			}
		}
		//返回所有的字段说明
		return annotationList;
	}
	//获取字段类型那一行的数据
	private List<String> getClasses(Row classRow) {
		//声明一个list，用来存放所有的字段类型
		List<String> classList = new ArrayList<String>();
		//判断这一行是否为空
		if (classRow != null) {
			//遍历这一行的所有单元格
			for (int i = 0; i < classRow.getLastCellNum(); i++) {
				//获取单元格
				Cell cell = classRow.getCell(i);
				//判断单元格是否为空
				if (cell != null) {
					//将单元格的内容存放到list中
					classList.add(cell.getStringCellValue());
				}
			}
		}
		//返回所有的字段类型
		return classList;
	}
	//获取标志位那一行的信息
	private Object[] getTypes(Row typeRow, int keyNum) {
		//声明一个List，用来存放所有要放入生成的client端的xml文件中的数据位置信息
		List<Integer> cnums = new ArrayList<Integer>();
		//声明一个List，用来存放所有要放入生成的server端的xml文件中的数据位置信息
		List<Integer> snums = new ArrayList<Integer>();
		//声明一个boolean值，用来判断是否要生成client端的xml文件
		boolean cfal = false;
		//声明一个boolean值，用来判断是否要生成server端的xml文件
		boolean sfal = false;
		//声明一个object数字，用来存放这一行所要返回的信息
		Object[] obj = new Object[4];
		//判断这一行是否为空
		if (typeRow != null) {
			//遍历这一行的所有单元格
			for (int i = 0; i < typeRow.getLastCellNum(); i++) {
				//获取单元格
				Cell cell = typeRow.getCell(i);
				//判断单元格是否为空
				if (cell != null) {
					//判断单元格的内容是否为C，为C表示这一列要放入到生成的client端要使用的xml文件中
					if (cell.getStringCellValue().equals("C")) {
						//添加单元格位置到cnums中
						cnums.add(i);
						//判断是否是key列，如果是，表示要生成client端使用的xml文件
						if (keyNum == i) {
							//将判断是否生成client端使用的xml文件的标志设为true
							cfal = true;
						}
					//判断单元格的内容是否为S，为S表示这一列要放入到生成的server端要使用的xml文件中
					} else if (cell.getStringCellValue().equals("S")) {
						//添加单元格位置到snums中
						snums.add(i);
						//判断是否是key列，如果是，表示要生成server端使用的xml文件
						if (keyNum == i) {
							//将判断是否生成server端使用的xml文件的标志设为true
							sfal = true;
						}
					//判断单元格的内容是否为A，为A表示这一列既要放入到生成的client端使用的xml文件中，又要放入到生成的server端要使用的xml文件中
					} else if (cell.getStringCellValue().equals("A")) {
						//添加单元格位置到cnums中
						cnums.add(i);
						//添加单元格位置到snums中
						snums.add(i);
						//判断是否是key列，如果是，表示要生成client端和server端使用的xml文件
						if (keyNum == i) {
							//将判断是否生成client端使用的xml文件的标志设为true
							cfal = true;
							//将判断是否生成server端使用的xml文件的标志设为true
							sfal = true;
						}
					}
				}
			}
		}
		//将要生成client端xml文件的位置信息放入到obj[0]
		obj[0] = cnums;
		//将要生成server端xml文件的位置信息放入到obj[1]
		obj[1] = snums;
		//将判断是否要生成client端xml文件的标志放入到obj[2]
		obj[2] = cfal;
		//将判断是否要生成server端xml文件的标志放入到obj[3]
		obj[3] = sfal;
		//返回这一行所有的信息
		return obj;
	}
	//获取要打印到xml文件中的内容
	private String getOutputStr(List<String> nameList, List<String> annotationList, List<String> classList,
			String firstName, Sheet sheet, List<Integer> nums,String xmlDeclaration, boolean isServer) {
		//声明一个StringBuilder，用来存放要打印到xml文件中的内容
		StringBuilder builder = new StringBuilder("");
		//向builder中放入配置声明，包括版本号和编码类型
		builder.append(xmlDeclaration +" \n");
		//向builder中放入注释的开始符号
		builder.append("<!-- ");
		//向builder中放入换行符
		builder.append("\n");
		//遍历位置信息
		for (Integer num : nums) {
			//将该位置的字段说明放入到builder中
			builder.append(annotationList.get(num) + " ");
		}
		//向builder中放入换行符
		builder.append("\n");
		//遍历位置信息
		for (Integer num : nums) {
			//将该位置的字段类型放入到builder中
			builder.append(classList.get(num) + " ");
		}
		//向builder中放入换行符
		builder.append("\n");
		//向builder中放入注释结束符和换行符
		builder.append("--> \n");
		//向builder中放入标签开始符号
		builder.append("<");
		//向builder中放入标签名称
		builder.append(firstName);
		//向builder中放入标签结束符号
		builder.append(">");
		//向builder中放入换行符
		builder.append("\n");
		//遍历该sheet中从第四行开始的所有行，从第四行开始就是数据行了
		for (int i = 4; i <= sheet.getLastRowNum(); i++) {
			//获取某一行
			Row row = sheet.getRow(i);
			//判断这一行是否为空
			if (row != null) {
				//获取本行的第一个单元格
				Cell cell1 = row.getCell(0);
				//判断这个单元格是否为空，或者是空白的，或者是错误的，如果符合其中一种，我们就跳过此次循环
				if (cell1 == null || cell1.getCellType() == Cell.CELL_TYPE_BLANK
						|| cell1.getCellType() == Cell.CELL_TYPE_ERROR) {
					continue;
				}
				//这里是向builder中添加数据标签的开始符，我这里用的是小写字母l，大家也可以换掉
				builder.append("    <l ");
				//遍历所有要向xml中放入的数据列
				for (int j = 0; j < nums.size(); j++) {
					//获取这一行中的某一列的单元格
					Cell cell = row.getCell(nums.get(j));
					//判断单元格是否为空，或者空白，或者是错误的，如果符合其中一种，我们就跳过此次循环
					if (cell != null) {
						if (cell.getCellType() == Cell.CELL_TYPE_BLANK
								|| cell.getCellType() == Cell.CELL_TYPE_ERROR) {
							continue;
						}
						//判断生成的xml文件是否是server端使用的，如果是server端使用的，我们将他的属性名全部小写处理，
						//如果大家不想这样做，可以去掉这个判断，只留下else里面的那一行代码
						if (isServer) {
							//向builder中添加该列的属性名称，这里是对server端的处理，所以我进行了toLowerCase()处理
							builder.append(nameList.get(nums.get(j)).toLowerCase() + "=\"");
						} else {
							//向builder中添加该列的属性名称，这里是对client端的处理
							builder.append(nameList.get(nums.get(j)) + "=\"");
						}
						//判断该单元格是否是boolean类型
						if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							//当该单元格的类型为boolean类型时，我们调用单元格的getBooleanCellValue()这个方法获取它的值，然后将该值放入builder中
							builder.append(cell.getBooleanCellValue() + "\" ");
						//判断该单元格是否是公式类型
						} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
							//当该单元格的类型为公式类型时，我们调用单元格的getCellFormula()这个方法获取它的值，然后将该值放入builder中（这个获取到的是一个公式）
							builder.append(cell.getCellFormula() + "\" ");
						//判断该单元格是否是数值类型
						} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							//当该单元格的类型为数值类型时，我们调用单元格的getNumericCellValue()这个方法获取它的值，然后将该值放入builder中
							//这里因为我用到的数据都是整数型的，所有我将取到的值强转成了int，大家也可以去掉强转，就取原来的值就好
							builder.append((int) (cell.getNumericCellValue()) + "\" ");
						//判断该单元格的类型是否是字符串类型
						} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							//当该单元格的类型为字符串类型时，我们调用单元格的getStringCellValue()这个方法获取它的值，然后将该值放入builder中
							builder.append(cell.getStringCellValue() + "\" ");
						}
					}
				}
				//向builder中放入本条数据的结束符，以及换行符
				builder.append("/> \n");
			} else {
				continue;
			}
		}
		//向builder中放入结束标签的开始符
		builder.append("</");
		//向builder中放入结束标签的标签名
		builder.append(firstName);
		//向builder中放入结束标签的结束符
		builder.append(">");
		//返回我们要向xml中插入的内容
		return builder.toString();
	}
	//将内容outputStr插入到文件xmlFile中
	private void writer(File xmlFile, String outputStr) throws IOException {
		//判断该文件是否存在
		if (xmlFile.exists()) {
			//如果存在，删除该文件
			xmlFile.delete();
		}
		//创建该文件
		xmlFile.createNewFile();
		//实例化一个输出流，该输出流输出的目标对象是xmlFile，输出时的编码格式为utf-8，这里大家可以根据自己的实际情况作修改
		OutputStreamWriter writer = new OutputStreamWriter(new FileOutputStream(xmlFile), "utf-8");
		//将outputStr写入到文本中
		writer.write(outputStr);
		//关闭输出流
		writer.close();
	}
}
