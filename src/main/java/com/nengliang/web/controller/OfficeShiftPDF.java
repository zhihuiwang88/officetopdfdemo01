package com.nengliang.web.controller;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class OfficeShiftPDF {

	private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;
	
	
	/**
	 * 使用JDK1.8
	 * jacob
	 * 实现Office转换为PDF
	 * @since20190708
	 * @author Dell
	 * @param args
	 * @throws Exception 
	 * srcFilePath : 要转换的文档位置
	 * pdfFilePath : 转换成功后文档的位置
	 */

	// doc转pdf
	public static String word2PDF(String srcFilePath, String pdfFilePath) throws Exception {
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			ComThread.InitSTA();
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", false);
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// 是否只读
			Object[] obj = new Object[] { srcFilePath, new Variant(false), new Variant(false), 
					new Variant(false), new Variant("pwd") };
			doc = Dispatch.invoke(docs, "Open", Dispatch.Method, obj, new int[1]).toDispatch();
			Dispatch.put(doc, "RemovePersonalInformation", false);
			// word保存为pdf格式宏，值为17
			Dispatch.call(doc, "ExportAsFixedFormat", pdfFilePath, WORD_TO_PDF_OPERAND); 

		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		} finally {
			if (doc != null) {
				Dispatch.call(doc, "Close", false);
			}
			if (app != null) {
				app.invoke("Quit", 0);
			}
			ComThread.Release();
		}
		return "doc2pdf";

	}

	// ppt转pdf
	public static String ppt2PDF(String srcFilePath, String pdfFilePath) throws Exception {
		ActiveXComponent app = null;
		Dispatch ppt = null;
		try {
			ComThread.InitSTA();
			app = new ActiveXComponent("PowerPoint.Application");
			Dispatch ppts = app.getProperty("Presentations").toDispatch();

			/*
			 * call param 4: ReadOnly param 5: Untitled指定文件是否有标题 param 6:
			 * WithWindow指定文件是否可见
			 */
			ppt = Dispatch.call(ppts, "Open", srcFilePath, true, true, false).toDispatch();
			// ppSaveAsPDF为特定值32
			Dispatch.call(ppt, "SaveAs", pdfFilePath, PPT_TO_PDF_OPERAND); 

		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		} finally {
			if (ppt != null) {
				Dispatch.call(ppt, "Close");
			}
			if (app != null) {
				app.invoke("Quit");
			}
			ComThread.Release();
		}
		return "ppt2pdf";
	}

	// excel转Pdf
	public static String excel2PDF(String inFilePath, String outFilePath) throws Exception {
		ActiveXComponent ax = null;
		Dispatch excel = null;
		try {
			ComThread.InitSTA();
			ax = new ActiveXComponent("Excel.Application");
			ax.setProperty("Visible", new Variant(false));
			// 禁用宏
			ax.setProperty("AutomationSecurity", new Variant(3)); 
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();

			Object[] obj = new Object[] { inFilePath, new Variant(false), new Variant(false) };
			excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();

			// 转换格式 , PDF格式=0
			Object[] obj2 = new Object[] { new Variant(EXCEL_TO_PDF_OPERAND), 
					// 0=标准 (生成的PDF图片不会变模糊) ; 1=最小文件
					outFilePath, new Variant(0) 
			};
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, obj2, new int[1]);

		} catch (Exception es) {
			es.printStackTrace();
			throw es;
		} finally {
			if (excel != null) {
				Dispatch.call(excel, "Close", new Variant(false));
			}
			if (ax != null) {
				ax.invoke("Quit", new Variant[] {});
				ax = null;
			}
			ComThread.Release();
		}
		return "excel2Pdf";

	}
    
	public static void main(String[] args) throws Exception {
         //   filePath 桌面
		String filePath = "E:\\";
		String fileNamePath = filePath + "test.docx";
		String fileNewPath = filePath + "aa2.pdf";
		Long  startTime = System.currentTimeMillis();
		
        new OfficeShiftPDF().word2PDF(fileNamePath,fileNewPath);
        //new OfficeShiftPDF().excel2PDF(fileNamePath,fileNewPath);
        //new OfficeShiftPDF().ppt2PDF(fileNamePath,fileNewPath);
        Long  endTime = System.currentTimeMillis();
        System.out.println("office转换pdf耗时："+ (endTime - startTime) + "毫秒。");
        System.out.println();
        System.out.println("Office 转 PDF成功！");
        
    }
	
	
	
	
}
