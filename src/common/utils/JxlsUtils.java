package common.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxls.demo.guide.Employee;

import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

/**
 * 
 * @author klguang
 *
 */
public class JxlsUtils{
	
	//private static final String TEMPLATE_PATH="jxls-template";
	private static final String TEMPLATE_PATH="E:";
	
	public static void exportExcel(InputStream is, OutputStream os, Map<String, Object> model) throws IOException{
        Context context = new Context();
        if (model != null) {
            for (String key : model.keySet()) {
            	System.out.println(key+"-----"+model.get(key));
            	if (key =="employees") {
            		 List<Employee> employees = (List<Employee>) model.get(key);
            		 for (Employee employee : employees) {
						
            			 System.out.println(employee.getBirthDate());
					}
				}
                context.putVar(key, model.get(key));
            }
        }
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        Transformer transformer  = jxlsHelper.createTransformer(is, os);
        //获取配置
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator)transformer.getTransformationConfig().getExpressionEvaluator();
        //函数强制，自定义功能
        Map<String, Object> funcs = new HashMap<String, Object>();
        funcs.put("utils", new JxlsUtils());    //添加自定义功能
        evaluator.getJexlEngine().setFunctions(funcs);
        //必须要这个，否者表格函数统计会错乱
        jxlsHelper.setUseFastFormulaProcessor(false).processTemplate(context, transformer);
	}

    public static void exportExcel(File xls, File out, Map<String, Object> model) throws FileNotFoundException, IOException {
            exportExcel(new FileInputStream(xls), new FileOutputStream(out), model);
    }
    
    public static void exportExcel(String templateName, OutputStream os, Map<String, Object> model) throws FileNotFoundException, IOException {
    	File template = getTemplate(templateName);
    	if(template!=null){
    		System.out.println("template is not null！");
        	exportExcel(new FileInputStream(template), os, model);	
    	}
    }
    
    
    //获取jxls模版文件

    public static File getTemplate(String name){
//        String templatePath = JxlsUtils.class.getClassLoader().getResource(TEMPLATE_PATH).getPath();
        String templatePath = TEMPLATE_PATH;
        System.out.println(templatePath);
        File template = new File(templatePath, name);
        System.out.println(template);
        System.out.println(template.exists());
        if(template.exists()){
            return template;
        }
        return null;
    }	
	
    // 日期格式化
    public String dateFmt(Date date, String fmt) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(fmt);
            return dateFmt.format(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }
    
    // if判断
    public Object ifelse(boolean b, Object o1, Object o2) {
        return b ? o1 : o2;
    }
    
}
