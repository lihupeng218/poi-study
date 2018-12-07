package cn.jiyun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class POIDemo {

	/**
	 * 导出Excel文件，高版本，后缀是 .xlsx
	 * 读取高版本用的是POI中的XSSF
	 * 工作簿：workbook
	 * 工作表：sheet
	 * 行：row
	 * 单元格：cell
	 * @throws Exception 
	 */
	@Test
	public void export() throws Exception{
		
		//创建一个工作簿
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		//创建一个工作表
		XSSFSheet sheet = workbook.createSheet("用户信息");
		
		//创建第一行，目的是为了写标题
		XSSFRow firstRow = sheet.createRow(0);
		
		//创建id的单元格
		XSSFCell idCell = firstRow.createCell(0);
		idCell.setCellValue("编号");
		
		//创建name的单元格
		XSSFCell nameCell = firstRow.createCell(1);
		nameCell.setCellValue("姓名");
		
		//创建gender的单元格
		XSSFCell genderCell = firstRow.createCell(2);
		genderCell.setCellValue("性别");
		
		//创建age的单元格
		XSSFCell ageCell = firstRow.createCell(3);
		ageCell.setCellValue("年龄");
		
		//从数据库查数据，这里做演示是写死的数据
		List<User> list = new ArrayList<>();
		list.add(new User(1, "张三", "男", 20));
		list.add(new User(2, "李四", "男", 22));
		list.add(new User(3, "小美", "女", 30));
		list.add(new User(4, "小花", "女", 20));
		list.add(new User(5, "刘能", "男", 20));
		list.add(new User(6, "小圆", "男", 20));
		
		//循环向第二行之后写数据
		for (int i = 0; i < list.size(); i++) {
			
			User user = list.get(i);
			
			//创建行
			XSSFRow row = sheet.createRow(i + 1);
			
			XSSFCell cell1 = row.createCell(0);
			cell1.setCellValue(user.getId());
			
			XSSFCell cell2 = row.createCell(1);
			cell2.setCellValue(user.getName());
			
			XSSFCell cell3 = row.createCell(2);
			cell3.setCellValue(user.getGender());
			
			XSSFCell cell4 = row.createCell(3);
			cell4.setCellValue(user.getAge());
		}
		
		//生成Excel文件
		//创建字节输出流对象
		FileOutputStream out = new FileOutputStream(new File("D:\\export.xlsx"));
		//生成Excel文件
		workbook.write(out);
		
		//刷新
		out.flush();
		//释放资源
		out.close();
	}
	
	/**
	 * 导入Excel文件，高版本，后缀是 .xlsx
	 * 读取高版本用的是POI中的XSSF
	 * 工作簿：workbook
	 * 工作表：sheet
	 * 行：row
	 * 单元格：cell
	 * @throws Exception 
	 * @throws Exception 
	 */
	@Test
	public void importExcel() throws Exception{
		//获取到一个工作簿
		FileInputStream in = new FileInputStream("D:\\import.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		
		//获取工作表
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//获取最后一行的行号
		int lastRowNum = sheet.getLastRowNum();
		
		List<User> list = new ArrayList<>();
		//循环读取Excel，第二行之后的数据
		for (int i = 1; i <= lastRowNum; i++) {
			
			//获取一个row
			XSSFRow row = sheet.getRow(i);
			
			//获取第一个单元格的编号值
			double id = row.getCell(0).getNumericCellValue();
			String name = row.getCell(1).getStringCellValue();
			String gender = row.getCell(2).getStringCellValue();
			double age = row.getCell(3).getNumericCellValue();

			User user = new User((int)id, name, gender, (int)age);
			list.add(user);
		}
		
		//将list中的数据插入到数据库中
		for (User user : list) {
			System.out.println(user);
		}
		
		//关闭资源
		in.close();
	}
}
