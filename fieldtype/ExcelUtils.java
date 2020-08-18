package com.jeesite.common.utils.excel.fieldtype;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.jeesite.modules.audit.enums.ExcelReportFileModelPath;

import java.io.*;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;

/**
 * @description: 基于apache POI操作Excel工具类
 * @author: Mr.Luke
 * @create: 2019-07-16 15:20
 * @Version V1.0
 */
public  class ExcelUtils {
	
	private static Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    /***
     * @Author: Mr.Luke
     * @Description: 将数据导入Excel
     * @Date: 16:30 2019/7/16
     * @Param: [excelReportFilePath, savePath,sheetName, list]
     * @return: void
     */
    public static void importXlsx(ExcelReportFileModelPath excelReportFileModelPath,String savePath, String sheetName, List list,int[] cellLine){
        
    	XSSFWorkbook workbook=importXlsxFile(excelReportFileModelPath,sheetName,list,cellLine,0);
        //输出到磁盘中
        FileOutputStream fos=null;
		try {
			fos = new FileOutputStream(new File(savePath));
			workbook.write(fos);
		} catch (FileNotFoundException e) {
			logger.error("输出文件不存在",e);
		} catch (IOException e) {
			logger.error("写入失败",e);
		}finally {
			 try {
				 if(workbook!=null){
					 workbook.close();
				 }
				 if(fos!=null){
					 fos.close(); 
				 }
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}        
    }
    

    /***
     * @Author: Mr.Luke
     * @Description: 将数据导入Excel
     * @Date: 16:30 2019/7/16
     * @Param: [excelReportFilePath, savePath,sheetName, list]
     * @return: void
     */
    public static void importXlsx(ExcelReportFileModelPath excelReportFileModelPath,OutputStream outputStream, String sheetName, List list,int cellLine){
    	XSSFWorkbook workbook=importXlsxFile(excelReportFileModelPath,sheetName,list,null,cellLine);
		try {
			workbook.write(outputStream);
		}catch (IOException e) {
			logger.error("写入失败",e);
		}finally {
			 try {
				 if(workbook!=null){
					 workbook.close();
				 }
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}
    }
    
    
    /***
     * @Author: Mr.Luke
     * @Description: 将数据导入Excel
     * @Date: 16:30 2019/7/16
     * @Param: [excelReportFilePath, savePath,sheetName, list]
     * @return: void
     */
    public static void importXlsx(ExcelReportFileModelPath excelReportFileModelPath,OutputStream outputStream, String sheetName, List list,int[] cellLine){
        
    	XSSFWorkbook workbook=importXlsxFile(excelReportFileModelPath,sheetName,list,cellLine,0);
		try {
			workbook.write(outputStream);
		}catch (IOException e) {
			logger.error("写入失败",e);
		}finally {
			 try {
				 if(workbook!=null){
					 workbook.close();
				 }
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}
       
        
    }
    
    
    
    private static XSSFWorkbook  importXlsxFile(ExcelReportFileModelPath excelReportFileModelPath , String sheetName, List list,int[] cellLine,int countLine){
    	//获得模板
    	InputStream in =  ExcelUtils.class.getClassLoader().getResourceAsStream("reportTemplate/"+excelReportFileModelPath.getCode());
    	
        //由输入流得到工作簿
        XSSFWorkbook workbook=null;
		try {
			workbook = new XSSFWorkbook(in);
		} catch (IOException e) {
			logger.error("模板文件导入失败",e);
		}finally {
			try {
				if(in!=null){
					in.close();
				}
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}
        
        if(sheetName==null){
        	sheetName="Sheet1";
        }
      //得到工作表
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet==null){
            sheet=workbook.createSheet(sheetName);
        }
        //获取新增行的行数
        int startRow=sheet.getLastRowNum()+1;
        int indexRow=sheet.getLastRowNum();

        for (Object model : list) {
            int r=++indexRow;
            XSSFRow row = sheet.createRow(r);
            //获取新增的列数
            try {
				writeRow(model, row);
			} catch (IllegalAccessException e) {
				logger.error("新增数据失败",e);
			}
        }
        //合并相同内容单元格
        if(cellLine==null){
        	 mergeCell(sheet,countLine,startRow,sheet.getLastRowNum());
        }else{
        	 mergeCell(sheet,cellLine,startRow,sheet.getLastRowNum());
        }
        return workbook;
    	
    }
    
    
    /**
     * 列表注入从第n行开始填入
     * @param excelReportFileModelPath
     * @param sheetName
     * @param list
     * @param cellLine
     * @param countLine
     * @return
     */
    private static XSSFWorkbook  importXlsxFile(String fileurl, String sheetName, List list,int startRow){
    	//获得模板
    	InputStream in =  ExcelUtils.class.getClassLoader().getResourceAsStream(fileurl);
        //由输入流得到工作簿
        XSSFWorkbook workbook=null;
		try {
			workbook = new XSSFWorkbook(in);
		} catch (IOException e) {
			logger.error("模板文件导入失败",e);
		}finally {
			try {
				if(in!=null){
					in.close();
				}
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}
        
        if(sheetName==null){
        	sheetName="Sheet1";
        }
      //得到工作表
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet==null){
            sheet=workbook.createSheet(sheetName);
        }
        //开始填充行
        int indexRow=startRow;
        for (Object model : list) {
            int r=++indexRow;
            XSSFRow row = sheet.createRow(r);
            //获取新增的列数
            try {
				writeRow(model, row);
			} catch (IllegalAccessException e) {
				logger.error("新增数据失败",e);
			}
        }
        return workbook;
    	
    }
    
    /**
     * 通过XSSFWorkbook 填充
     * @param sheetName
     * @param workbook
     * @param list
     * @param startRow
     * @return
     */
    private static XSSFWorkbook  importXlsxFile(String sheetName,XSSFWorkbook workbook, List list,int startRow){
        
        if(sheetName==null){
        	sheetName="Sheet1";
        }
      //得到工作表
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet==null){
            sheet=workbook.createSheet(sheetName);
        }
        //开始填充行
        int indexRow=startRow;
        for (Object model : list) {
            int r=++indexRow;
            XSSFRow row = sheet.createRow(r);
            //获取新增的列数
            try {
				writeRow(model, row);
			} catch (IllegalAccessException e) {
				logger.error("新增数据失败",e);
			}
        }
        return workbook;
    	
    }
    
    
    /***
     * @Author: Mr.Luke
     * @Description: 单独对象填入数据
     * @Date: 9:45 2019/7/23
     * @Param: [sheetName, file, cla]
     * @return: java.lang.Object
     * @throws IllegalAccessException 
     * @throws InstantiationException 
     */
    public static void importObjectDataFrom(OutputStream outputStream,ExcelReportFileModelPath excelReportFileModelPath , String sheetName, Object object) {
    	importObjectDataFrom(outputStream,"reportTemplate/"+excelReportFileModelPath.getCode(),sheetName,object);
    }

    
    
    
    /**
     *	设置文件URL 对象注入
     * @param outputStream
     * @param fileUrl
     * @param sheetName
     * @param object
     */
    public static void importObjectDataFrom(OutputStream outputStream,String fileUrl, String sheetName, Object object) {
    	//获得模板
    	InputStream in =  ExcelUtils.class.getClassLoader().getResourceAsStream(fileUrl);
    	
        //由输入流得到工作簿
        XSSFWorkbook workbook=null;
		try {
			workbook = new XSSFWorkbook(in);
		} catch (IOException e) {
			logger.error("模板文件导入失败",e);
		}finally {
			try {
				if(in!=null){
					in.close();
				}
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}
        
        if(sheetName==null){
        	sheetName="Sheet1";
        }
      //得到工作表
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet==null){
            sheet=workbook.createSheet(sheetName);
        }

        Class cla=object.getClass();
        
        Field[] fields = cla.getDeclaredFields();
        for(Field field : fields){
            boolean fieldHasAnno = field.isAnnotationPresent(ExcelCoordinate.class);

            if(fieldHasAnno){
                ExcelCoordinate fieldAnno = field.getAnnotation(ExcelCoordinate.class);
                //输出注解属性
                field.setAccessible(true);
                int row = fieldAnno.row();
                int col = fieldAnno.col();
                
                logger.debug(field.getName()+"\t "+row +" : "+col+"\t ");
                Row rows=sheet.getRow(row);
                Cell ce0 =  rows.getCell(col);
                Object value=null;
				try {
					value = field.get(object);
				} catch (IllegalArgumentException e) {
					logger.error("属性："+field.getName()+"获取数据失败。",e);
				} catch (IllegalAccessException e) {
					logger.error("属性："+field.getName()+"获取数据失败.",e);
				}
                if(value instanceof Date) {
                	SimpleDateFormat format=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            		ce0.setCellValue(format.format((Date)value));
                }else if(value==null){
                	ce0.setCellValue("");
                }else{
                    ce0.setCellValue(""+ value);
                }
                              
                logger.debug("value :"+value);
            
            }else if(field.isAnnotationPresent(ExcelListStart.class)){
            	ExcelListStart start = field.getAnnotation(ExcelListStart.class);
            	field.setAccessible(true);
            	try {
					importXlsxFile(sheetName,workbook,((List)field.get(object)),start.value());
				} catch (IllegalArgumentException e) {
					logger.error("属性："+field.getName()+"，貌似不是List对象.",e);
				} catch (IllegalAccessException e) {
					logger.error("属性："+field.getName()+"，貌似不是List对象.",e);
				}
            }
        }
        
        try {
			workbook.write(outputStream);
		}catch (IOException e) {
			logger.error("写入失败",e);
		}finally {
			 try {
				 if(workbook!=null){
					 workbook.close();
				 }
			} catch (IOException e) {
				logger.error("关闭失败",e);
			}
		}
    }

    private static void writeRow(Object model,XSSFRow row) throws IllegalAccessException {
        Field[] fields = model.getClass().getDeclaredFields();
        for(Field field : fields){
            boolean fieldHasAnno = field.isAnnotationPresent(ExcelCell.class);
            field.setAccessible(true);       
            //获取属性值
            Object value = field.get(model);
            if (value==null){
                continue;
            }
            if(fieldHasAnno&&isBaseType(value)){
                ExcelCell fieldAnno = field.getAnnotation(ExcelCell.class);
                //输出注解属性
                int index = fieldAnno.value();
                XSSFCell ce0 = row.createCell(index);
                if(value instanceof Date) {
                    SimpleDateFormat format=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    ce0.setCellValue(format.format((Date)value));
                }else{
                    ce0.setCellValue(""+ value);
                }
            }else if (!isBaseType(value)){
                writeRow(value,row);
            }
        }
    }

    /**
     * 判断object是否为基本类型 和字符串类型 和日期类型
     * @param object
     * @return
     */
    public static boolean isBaseType(Object object) {
        if (object==null){
            return false;
        }
        Class<? extends Object> className = object.getClass();
        if (className.equals(java.lang.Integer.class) ||
                className.equals(java.lang.Byte.class) ||
                className.equals(java.lang.Long.class) ||
                className.equals(java.lang.Double.class) ||
                className.equals(java.lang.Float.class) ||
                className.equals(java.lang.Character.class) ||
                className.equals(java.lang.Short.class) ||
                className.equals(java.lang.Boolean.class)||
                className.equals(String.class)||
                className.equals(Date.class)
        ) {
            return true;
        }
        return false;
    }


    /***
     * @Author: Mr.Luke
     * @Description: 合并内容相同待的单元格
     * @Date: 13:17 2019/7/17
     * @Param: [sheet, cellLine, startRow, endRow]
     * @return: void
     */
    private static void mergeCell(XSSFSheet sheet, int cellLine, int startRow, int endRow){
        for (int i=0;i<=cellLine;i++){
            mergeOneCell(sheet,i,startRow,endRow);
        }
    }
    
    /***
     * @Author: Mr.Luke
     * @Description: 合并内容相同待的单元格
     * @Date: 13:17 2019/7/17
     * @Param: [sheet, cellLine, startRow, endRow]
     * @return: void
     */
    private static void mergeCell(XSSFSheet sheet, int[] cellLine, int startRow, int endRow){
        for (int i=0;i<cellLine.length;i++){
            mergeOneCell(sheet,cellLine[i],startRow,endRow);
        }
    }



    /***
     * @Author: Mr.Luke
     * @Description: 合并一列的相同内容的单元格 
     * @Date: 13:16 2019/7/17
     * @Param: [sheet, cellLine, startRow, endRow]
     * @return: void
     */
    private static void mergeOneCell(XSSFSheet sheet, int cellLine, int startRow, int endRow){
        int s=startRow;
        String cellValue = getCellValue(sheet.getRow(s).getCell(cellLine));

        for (int i=s+1;i<=endRow+1;i++){
            Row row=sheet.getRow(i);
            //到了最后一行
            if (row==null){
                if(s!=i-1) {
                    try{
                		sheet.addMergedRegion(new CellRangeAddress(s, i-1, cellLine, cellLine));
                	}catch(IllegalStateException e){
                		logger.info("合并单元格错误："+s+","+(i-1)+","+cellLine+","+e.getMessage());
                	}
                }
                break;
            }
            String value = getCellValue(row.getCell(cellLine));
            //遇到空的单元格，或者和上一个内容不相符的单元格，往下移动
            if(value.equals("")||!cellValue.equals(value)) {
                //新的一行内容和之前的不一样，合并之前的单元格
                if(s!=i-1) {
                	try{
                		sheet.addMergedRegion(new CellRangeAddress(s, i-1, cellLine, cellLine));
                	}catch(IllegalStateException e){
                		logger.info("合并单元格错误："+s+","+(i-1)+","+cellLine+""+e.getMessage());
                	}
                }
                cellValue = value;
                s = i ;
            }
        }
    }


    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    private static  String getCellValue(Cell cell){
        if(cell == null){ return "";}
        //设置文本类型单元格获取
        
        DecimalFormat df = new DecimalFormat("0");  
		if(cell.getCellType()==Cell.CELL_TYPE_BOOLEAN){
			return  String.valueOf(cell.getBooleanCellValue());
		} else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
			Double value=cell.getNumericCellValue();
			if(value-value.intValue()==0){
				return df.format(value);  
			}else{
				return String.valueOf(value);
			}
		}else if(cell.getCellType()==Cell.CELL_TYPE_FORMULA){
			FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			CellValue cellValue = evaluator.evaluate(cell);
		      switch (cellValue.getCellType()) {
		        case Cell.CELL_TYPE_STRING:
		            return cellValue.getStringValue();
		        case Cell.CELL_TYPE_NUMERIC:
		        	Double value=cellValue.getNumberValue();
					if(value-value.intValue()==0){
						return df.format(value);  
					}else{
						df = new DecimalFormat("#.000");
						return df.format(value);
					}
		        default:
		            return "";
		        }
		}else{
			return  cell.getStringCellValue();  
		}
    }
    
  

    private static Workbook readExcel(File file) throws IllegalAccessException {
        // 根据后缀名称判断excel的版本
        String fileName = file.getName();
        String extName = fileName.substring(fileName.lastIndexOf(".") + 1);
        Workbook wb = null;
        if ("xls".equals(extName)) {
            try {
                wb = new HSSFWorkbook(new FileInputStream(file));
            } catch (IOException e) {
            	logger.error("读取文件错误",e);
            }

        } else if ("xlsx".equals(extName)) {
            try {
                wb = new XSSFWorkbook(new FileInputStream(file));
            } catch (IOException e) {
            	logger.error("读取文件错误",e);
            }

        } else {
            // 无效后缀名称，这里之能保证excel的后缀名称，不能保证文件类型正确，不过没关系，在创建Workbook的时候会校验文件格式
            throw new IllegalArgumentException("Invalid excel version");
        }
        return wb;
    }

    /***
     * @Author: Mr.Luke
     * @Description: 获取list表格 
     * @Date: 10:23 2019/7/23
     * @Param: [sheetName, file, cla, startRow, endRow]
     * @return: java.util.List
     * @throws IllegalAccessException 
     * @throws InstantiationException 
     */
    public static List getListData(String sheetName, File file, Class cla,int startRow,int endRow) throws InstantiationException, IllegalAccessException   {
        List list=new ArrayList();
        Workbook wb = null;
        try {
            wb=readExcel(file);
        } catch (IllegalAccessException e) {
        	logger.error("读取文件错误",e);
        }

        Sheet sheet=wb.getSheet(sheetName);
        for (int i=startRow;i<=endRow;i++){
            Field[] fields = cla.getDeclaredFields();
            Row row=sheet.getRow(i);
            if(row==null){
            	continue;
            }
            Object model=cla.newInstance();
            list.add(model);
            for(Field field : fields) {
                boolean fieldHasAnno = field.isAnnotationPresent(ExcelCell.class);
                field.setAccessible(true);
                if(fieldHasAnno){
                    ExcelCell fieldAnno = field.getAnnotation(ExcelCell.class);
                    //输出注解属性
                    int index = fieldAnno.value();
                   
                    String value=getCellValue(row.getCell(index));
                    try {
						setFieldValue(field,value,model);
					} catch (IllegalAccessException e) {
						logger.error("属性赋值错误",e);
					}
                }
            }
        }
        return list;
    }

    /***
     * @Author: Mr.Luke
     * @Description: 获取list表格 
     * @Date: 10:23 2019/7/23
     * @Param: [sheetName, file, cla, startRow, endRow]
     * @return: java.util.List
     * @throws IllegalAccessException 
     * @throws InstantiationException 
     */
    public static List getListDataHasLastRamaks(String sheetName, File file, Class cla,int startRow,int lastRamaksRow) throws InstantiationException, IllegalAccessException   {
        List list=new ArrayList();
        Workbook wb = null;
        try {
            wb=readExcel(file);
        } catch (IllegalAccessException e) {
        	logger.error("读取文件错误",e);
        }

        Sheet sheet=wb.getSheet(sheetName);
        lastRamaksRow=sheet.getLastRowNum()-lastRamaksRow;
        for (int i=startRow;i<=lastRamaksRow;i++){
            Field[] fields = cla.getDeclaredFields();
            Row row=sheet.getRow(i);
            if(row==null){
            	continue;
            }
            Object model=cla.newInstance();
            list.add(model);
            for(Field field : fields) {
                boolean fieldHasAnno = field.isAnnotationPresent(ExcelCell.class);
                field.setAccessible(true);
                if(fieldHasAnno){
                    ExcelCell fieldAnno = field.getAnnotation(ExcelCell.class);
                    //输出注解属性
                    int index = fieldAnno.value();
                
                    String value=getCellValue(row.getCell(index));
                    try {
						setFieldValue(field,value,model);
					} catch (IllegalAccessException e) {
						logger.error("属性赋值错误",e);
					}
                }
            }
        }
        return list;
    }
    

    
    /***
     * @Author: Mr.Luke
     * @Description: 获取单独对象数据
     * @Date: 9:45 2019/7/23
     * @Param: [sheetName, file, cla]
     * @return: java.lang.Object
     * @throws IllegalAccessException 
     * @throws InstantiationException 
     */
    public static Object getObjectDataFrom(String sheetName, File file, Class cla) throws InstantiationException, IllegalAccessException   {
        Workbook wb = null;
        try {
            wb=readExcel(file);
        } catch (IllegalAccessException e) {
        	logger.error("读取文件错误",e);
        }

        Sheet sheet=wb.getSheet(sheetName);
        Object model=cla.newInstance();
        Field[] fields = cla.getDeclaredFields();
        for(Field field : fields){
            boolean fieldHasAnno = field.isAnnotationPresent(ExcelCoordinate.class);

            if(fieldHasAnno){
                ExcelCoordinate fieldAnno = field.getAnnotation(ExcelCoordinate.class);
                //输出注解属性
                field.setAccessible(true);
                int row = fieldAnno.row();
                int col = fieldAnno.col();
                
                logger.debug(field.getName()+"\t "+row +" : "+col+"\t ");
                Row rows=sheet.getRow(row);
                
                String cellValue = getCellValue(rows.getCell(col));
                
                logger.debug("value :"+cellValue);
            
                try {
					setFieldValue(field,cellValue,model);
				} catch (IllegalAccessException e) {
					logger.error("属性赋值错误",e);
				}
            }
        }
        return model;
    }

    /***
     * @Author: Mr.Luke
     * @Description: 给对象属性赋值
     * @Date: 9:39 2019/7/23
     * @Param: [field, value, model]
     * @return: void
     */
    private static void setFieldValue(Field field,String value,Object model) throws IllegalAccessException {

        if(field.getType().equals(Date.class)) {
            
            Calendar calendar = new GregorianCalendar(1900,0,-1,0,0,0);
            calendar.set(Calendar.MILLISECOND, 0);
            try{
            	calendar.add(Calendar.DATE, Integer.parseInt(value));
            	field.set(model,calendar.getTime());
            }catch(Exception e){
            	SimpleDateFormat format=new SimpleDateFormat("yyyy/MM/dd");
            	try {
    				field.set(model,format.parse(value));   				
    			}catch (ParseException e2) {
    				logger.error("时间格式错误",e2);
    			}
            }

        }else if (field.getType().equals(Integer.class)){
            field.setInt(model,Integer.parseInt(value));
        }else if (field.getType().equals(Double.class)){
            field.setDouble(model,Double.parseDouble(value));
        }else if (field.getType().equals(Float.class)){
            field.setFloat(model,Float.parseFloat(value));
        }else if (field.getType().equals(String.class)){
            field.set(model,value);
        }
    }





}
