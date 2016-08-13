package com.lx.ui;

import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.lx.tools.ExcelReader;
import com.lx.tools.ExcelRow;

public class MainFrame extends JFrame{
	JPanel panel_1, panel_2, panel_3;
	JLabel label_1, label_2;
	JTextField text_1, text_2;
	JButton chooseBtn_1, chooseBtn_2, compareBtn;
	
	MainFrame(){
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		
		setLayout(new BoxLayout(getContentPane(), BoxLayout.Y_AXIS));
		
		panel_1 = new JPanel();
		this.add(panel_1);
		panel_1.setLayout(new FlowLayout());
		label_1 = new JLabel("文件1:");
		panel_1.add(label_1);
		text_1 = new JTextField(20);
		panel_1.add(text_1);
		chooseBtn_1 = new JButton("选择文件");
		panel_1.add(chooseBtn_1);
		chooseBtn_1.addActionListener(new ActionListener(){

			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				JFileChooser chooser = new JFileChooser("./");
				FileNameExtensionFilter filter = new FileNameExtensionFilter("excel文件 xls xlsx", "xls", "xlsx");
				chooser.setFileFilter(filter);
				int returnVal = chooser.showOpenDialog(MainFrame.this);
				if(returnVal == JFileChooser.APPROVE_OPTION) {
					MainFrame.this.text_1.setText(chooser.getSelectedFile().getAbsolutePath());
				}
			}
			
		});
		
		panel_2 = new JPanel();
		this.add(panel_2);
		panel_2.setLayout(new FlowLayout());
		label_2 = new JLabel("文件2:");
		panel_2.add(label_2);
		text_2 = new JTextField(20);
		panel_2.add(text_2);
		chooseBtn_2 = new JButton("选择文件");
		panel_2.add(chooseBtn_2);
		chooseBtn_2.addActionListener(new ActionListener(){

			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				JFileChooser chooser = new JFileChooser("./");
				FileNameExtensionFilter filter = new FileNameExtensionFilter("excel文件 xls xlsx", "xls", "xlsx");
				chooser.setFileFilter(filter);
				int returnVal = chooser.showOpenDialog(MainFrame.this);
				if(returnVal == JFileChooser.APPROVE_OPTION) {
					MainFrame.this.text_2.setText(chooser.getSelectedFile().getAbsolutePath());
				}
			}
			
		});
		
		panel_3 = new JPanel();
		this.add(panel_3);
		compareBtn = new JButton("开始对比");
		panel_3.add(compareBtn);
		compareBtn.addActionListener(new ActionListener(){

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser fc = new JFileChooser();
			    fc.setSelectedFile(new File(""));
		        int returnVal = fc.showSaveDialog(MainFrame.this);
		        if (returnVal == JFileChooser.APPROVE_OPTION) {
		            File file = fc.getSelectedFile();
		            //This is where a real application would save the file.
		            if(file != null){
		            	if(verifyDuplicate(file.getAbsolutePath() + ".xls")){
		            		int jop = JOptionPane.showConfirmDialog(MainFrame.this, "该文件已存在是否覆盖？", "提示", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
		            		if(jop == 0){
		            			exportExcel(MainFrame.this.text_1.getText().toString(), MainFrame.this.text_2.getText().toString(), file.getAbsolutePath() + ".xls");
		            			System.out.println("Saving: " + file.getAbsolutePath() + ".");
		            		} else {
		            			System.out.println("Not Saving: " + file.getAbsolutePath() + ".");
		            		}
		            	} else {
		            		exportExcel(MainFrame.this.text_1.getText().toString(), MainFrame.this.text_2.getText().toString(), file.getAbsolutePath() + ".xls");
		            		System.out.println("Saving: " + file.getAbsolutePath() + ".");
		            	}
		            } else {
//		            	JOptionPane.showMessageDialog(this.huodanPanel, "文件名不能为空！", "文件名为空", JOptionPane.PLAIN_MESSAGE);
		            }
		            
		        } else {
		            System.out.println("Save command cancelled by user.");
		        }
			}
			
			private boolean verifyDuplicate(String file_absolutepath){
				File file = new File(file_absolutepath);
				boolean exists = file.exists();
				return exists;
			}
			
			private void exportExcel(String file1_absolutepath, String file2_absolutepath, String file_absolutepath){
				ArrayList<ExcelRow> arraylist1 = new ExcelReader(file1_absolutepath).getAl();
				ArrayList<ExcelRow> arraylist2 = new ExcelReader(file2_absolutepath).getAl();
				
				HSSFWorkbook wb = new HSSFWorkbook();
				HSSFSheet sheet = wb.createSheet("油价对账");
				HSSFCellStyle cellStyle = wb.createCellStyle();
				HSSFDataFormat format= wb.createDataFormat();
				cellStyle.setDataFormat(format.getFormat("m月d日"));
				HSSFCellStyle cellStyle2 = wb.createCellStyle();
			    Font font = wb.createFont();
			    font.setColor(HSSFColor.RED.index);    //绿字
			    cellStyle2.setFont(font);
				int pipeishu = 0;
				for(int i=0; i<arraylist1.size(); i++){
					System.out.println(i);
					HSSFRow row = sheet.createRow(i);
					HSSFCell cell = row.createCell(0);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(arraylist1.get(i).getRiqi());
					row.createCell(1).setCellValue(arraylist1.get(i).getChepaihao());
					row.createCell(2).setCellValue(arraylist1.get(i).getShengshu());
					row.createCell(3).setCellValue(arraylist1.get(i).getYoujia());
					
					for(int j=0; j<arraylist2.size(); j++){
						if(arraylist2.get(j).getShifoupeidui() == 0
								&& arraylist1.get(i).getRiqi().equals(arraylist2.get(j).getRiqi()) 
								&& arraylist1.get(i).getChepaihao().equals(arraylist2.get(j).getChepaihao())
								){
							arraylist1.get(i).setShifoupeidui(1);//找到匹配的设置为1
							arraylist2.get(j).setShifoupeidui(1);
//							SimpleDateFormat sd=new SimpleDateFormat("yyyy年MM月DD日");
							cell = row.createCell(4);
							cell.setCellStyle(cellStyle);
							cell.setCellValue(arraylist2.get(j).getRiqi());
//							row.createCell(4).setCellValue(arraylist2.get(j).getRiqi());
							row.createCell(5).setCellValue(arraylist2.get(j).getChepaihao());
							row.createCell(6).setCellValue(arraylist2.get(j).getShengshu());
							if(arraylist1.get(i).getYoujia() == arraylist2.get(j).getYoujia()){
								row.createCell(7).setCellValue(arraylist2.get(j).getYoujia());
							} else {
								cell = row.createCell(7);
								cell.setCellStyle(cellStyle2);
								cell.setCellValue(arraylist2.get(j).getYoujia());
//								row.createCell(7).setCellValue(arraylist2.get(j).getYoujia());
							}
							pipeishu++;
						}
					}
				}
				int n=0;
				if(pipeishu < arraylist2.size()){
					for(int i=0; i<arraylist2.size(); i++){
						if(arraylist2.get(i).getShifoupeidui() == 0){
							n++;
							HSSFRow row = sheet.createRow(arraylist1.size()-1+n);
							row.createCell(0).setCellValue("");
							row.createCell(1).setCellValue("");
							row.createCell(2).setCellValue("");
							row.createCell(3).setCellValue("");
							HSSFCell cell = row.createCell(0);
							cell = row.createCell(4);
							cell.setCellStyle(cellStyle);
							cell.setCellValue(arraylist2.get(i).getRiqi());
//							row.createCell(4).setCellValue(arraylist2.get(i).getRiqi());
							row.createCell(5).setCellValue(arraylist2.get(i).getChepaihao());
							row.createCell(6).setCellValue(arraylist2.get(i).getShengshu());
							row.createCell(7).setCellValue(arraylist2.get(i).getYoujia());
						}
					}
				}
				
				outputExcel(wb, file_absolutepath);
			}
			
			private void outputExcel(HSSFWorkbook wb, String file_absolutepath){
				FileOutputStream fileOut;
				try {
					fileOut = new FileOutputStream(file_absolutepath);
					try {
						wb.write(fileOut);
						fileOut.close();
						JOptionPane.showMessageDialog(MainFrame.this, "对比成功！", "对比成功", JOptionPane.PLAIN_MESSAGE);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}

		});
		
		this.pack();
		setDefaultCloseOperation(EXIT_ON_CLOSE);
		this.setVisible(true);
	}
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new MainFrame();
	}
}
