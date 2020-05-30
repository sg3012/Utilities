package com.utils.excel;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Locale;
import java.util.Scanner;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelops {
	public FileInputStream ip = null;
	public FileOutputStream out = null;
	public File file = null;
	public String path;
	private Scanner sc = null;
	private String Sheetname = null;
	XSSFWorkbook wb = null;
	XSSFSheet sheet = null;
	XSSFRow row = null;
	XSSFCell cell = null;
	XSSFCellStyle style = null;
	boolean flag = false;
	private String cellvalue = null;

	public Excelops(String path) {
		// TODO Auto-generated constructor stub
		/*
		 * THIS METHOD LOADS AN EXCEL WORKBOOK IN THE SCRIPT TAKING PATH TO THE WORKBOOK
		 * Arguments - path - takes string value
		 */
		try {
			// 1. Connect the file path to the code:
			file = new File(path);

			// 2. Load the actual file in the form of Bytes in the code :
			ip = new FileInputStream(file);
			// 3. Load the file in xlsx format:
			wb = new XSSFWorkbook(ip);

		}

		catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void createblankworkbook(String path, String sheetname) {
		/*
		 * THIS METHOD CREATES A BLANK WORKBOOK HAVING ONLY ONE SHEET WITH i ROWS and j
		 * COLUMNS Arguments - path - takes string value sheetname - takes string value
		 * 
		 */
		try {

			// 1. Connect the file path to the code:
			file = new File(path);

			// 2. Create a new XSSFWorkBook by passing the file class reference:
			wb = new XSSFWorkbook(file);
			// 3. Create new sheet for:

			sheet = wb.createSheet(sheetname);

			// 3. Load the workbook in output stream :
			out = new FileOutputStream(file);

			// 4. Write the workbook in the path :
			wb.write(out);
			out.close();
		}

		catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void createblankworkbook(String path, int numberofsheets) {
		// THIS METHOD CREATES A BLANK WORKBOOK WITH MULTIPLE SHEETS HAVING i ROWS and j
		// COLUMNS :
		/*
		 * path - Take string value. numberofsheets - Takes integer values
		 */
		int i, j;

		String sheetelement1, sheetelement2;
		try {
			int no = Integer.parseInt(Integer.toString(numberofsheets));

			if (numberofsheets == 0 || Integer.toString(numberofsheets) == null
					|| Integer.toString(numberofsheets).startsWith("-")) {
				System.out.println("Invalid Number of sheets-Please Enter correct number");
			}

			else {
				// 1. Connect the file path to the code:
				file = new File(path);

				// 2. Create a new XSSFWorkBook by passing the file class reference:
				wb = new XSSFWorkbook();
				// 3. Create multiple sheets in the workbook:
				sc = new Scanner(System.in);
				ArrayList<String> list = new ArrayList<String>();
				System.out.println("LIST: " + list);
				for (i = 0; i < numberofsheets; i++) {
					System.out.println("Enter sheet" + (i + 1) + " " + "name");
					Sheetname = sc.nextLine();
					if (Sheetname.startsWith("\\") || Sheetname.contains("/") || Sheetname.contains("?")
							|| Sheetname.contains("*") || Sheetname.contains("[") || Sheetname.contains("]")) {
						System.out.println("Invalid sheet Name-Please Try again");
						i = i - 1;

					} else if (Sheetname == null || Sheetname.contains(" ") || (Sheetname.length() <= 0)) {

						System.out.println("Sheetname cannot be blank-Please try again");
						i = i - 1;
					} else {
						list.add(Sheetname);
						if (list.size() == 1) {
							sheet = wb.createSheet(list.get(i).toString());
						} else if (list.size() > 1) {
							for (j = (i - 1); j >= 0; j--) {
								sheetelement1 = list.get(i).toString();
								sheetelement2 = list.get(j).toString();
								if (list.get(i).toString().equals(list.get(j).toString())) {
									flag = true;
								}
							}

							if (flag == true) {
								list.remove(i);
								i = i - 1;
							} else {

								sheet = wb.createSheet(list.get(i).toString());

							}
							flag = false;
						}

					}
				}
				sc.close();

				// 3. Load the workbook in output stream :
				out = new FileOutputStream(file);

				// 4. Write the workbook in the path :
				wb.write(out);
				out.close();
			}

		}

		catch (NumberFormatException e) {
			System.out.println("Invalid number of sheets added-Please enter correct number");
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public boolean isSheetexist(String sheetname) {
		/*
		 * THIS METHOD ADD A SHEET IN EXISTING WORKBOOK 
		 * sheetname - takes string value
		 */

		 int index = wb.getSheetIndex(sheetname);
		
		 if (index == -1) 
		   {
			
			 
			 index = wb.getSheetIndex(sheetname.toUpperCase());
			   
			     if (index == -1)
				    return false;
			         
			      else
				      return true;
		     } 
		 else
			return true;
	}

	
	public boolean addsheet(String sheetName)
	{
		/*
		 * THIS METHOD CREATES A BLANK SHEET IN EXISTING WORBOOK WITH i ROWS and j
		 * COLUMNS 
		 * Arguments -  sheetname - takes string value
		 * 
		 */
		try {

			sheet = wb.createSheet(sheetName);
			//Load the workbook in output stream :
			out = new FileOutputStream(file);
			// 4. Write the workbook in the path :
			wb.write(out);
			out.close();
		    }

		      catch (Exception e) {
			    e.printStackTrace();
			    return false; 
		      }
		
		return true;
	}
	


	public String getData(int sheetIndex, int rowIndex, int colIndex) {

		/*
		 * THIS METHOD RETURNS THE VALUE AT A PARTICULAR CELL IN A SHEET.
		 * 
		 * Arguments - sheetIndex - takes integer value rowIndex - takes integer value
		 * colIndex -takes integer value
		 */

		try {
			if (!(sheetIndex > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return "";
			} else {
				// 1.
				/*
				 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
				 * will take the sheet number as integer value Index where 0 will indicate the
				 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
				 */
				try {
					sheet = wb.getSheetAt(sheetIndex - 1);

					if (sheet == null) {
						System.out.println("Sheet Doesn't Exist");
						return "";
					}
					if (!((rowIndex - 1) > 0)) {
						return "";
					} else if (!((colIndex - 1) >= 0)) {
						return "";
					}

					row = sheet.getRow(rowIndex - 1);

					if (row == null)
						return "";

					cell = row.getCell(colIndex - 1);
					if (cell == null)
						return "";

//           			      System.out.println("Cell Type: "+cell.getCellType());
					if (cell.getCellType().toString().equals("STRING")) {
						cellvalue = cell.getStringCellValue();
						return cellvalue;

					} else if (cell.getCellType().toString().equals("NUMERIC")
							|| cell.getCellType().toString().equals("FORMULA")) {
						cellvalue = cell.getRawValue();
						// System.out.println("Numeric cell value: "+cell.getNumericCellValue());
						if (DateUtil.isCellDateFormatted(cell)) {
							double d = cell.getNumericCellValue();
							Date date = DateUtil.getJavaDate(d);
							String pattern = "dd-MM-yyyy";
							SimpleDateFormat simpledateformat = new SimpleDateFormat(pattern, new Locale("en", "UK"));
							cellvalue = simpledateformat.format(date);
//                 				    System.out.println("date cell value: "+cellvalue);
						}
						return cellvalue;
					}

					else if (cell.getCellType().toString().equals("BLANK"))

						return "";

					else {
						cellvalue = String.valueOf(cell.getBooleanCellValue());
						return cellvalue;
					}

				}

				catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}

		}

		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return cellvalue;
	}

	public String getData(int sheetIndex, int rowIndex, String colname) {

		/*
		 * THIS METHOD RETURNS THE VALUE AT A PARTICULAR CELL IN A SHEET.
		 * 
		 * Arguments - sheetIndex - takes integer value rowIndex - takes integer value
		 * colname -takes string value
		 */
		int colIndex = -1;
		try {
			if (!(sheetIndex > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return "";
			}

			else {

				// 1.
				/*
				 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
				 * will take the sheet number as integer value Index where 0 will indicate the
				 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
				 */
				try {
					sheet = wb.getSheetAt(sheetIndex - 1);

					if (sheet == null) {
						System.out.println("Sheet Doesn't Exist");
						return "";
					}

					if (!((rowIndex - 1) > 0)) {
						return "";
					}

					row = sheet.getRow(0);

					for (int i = 0; i < row.getLastCellNum(); i++) {
						// System.out.println(row.getCell(i).getStringCellValue().trim());
						if (row.getCell(i).getStringCellValue().trim().equals(colname.trim()))
							colIndex = i;
					}

					if (colIndex == -1)
						return "";

					sheet = wb.getSheetAt(sheetIndex - 1);

					row = sheet.getRow(rowIndex - 1);

					if (row == null)
						return "";

					cell = row.getCell(colIndex);

					if (cell == null)
						return "";

//           			      System.out.println("Cell Type: "+cell.getCellType());
					if (cell.getCellType().toString().equals("STRING")) {
						cellvalue = cell.getStringCellValue();
						return cellvalue;

					}

					else if (cell.getCellType().toString().equals("NUMERIC")
							|| cell.getCellType().toString().equals("FORMULA"))

					{
						cellvalue = cell.getRawValue();
						// System.out.println("Numeric cell value: "+cell.getNumericCellValue());
						if (DateUtil.isCellDateFormatted(cell)) {
							double d = cell.getNumericCellValue();
							Date date = DateUtil.getJavaDate(d);
							String pattern = "dd-MM-yyyy";
							SimpleDateFormat simpledateformat = new SimpleDateFormat(pattern, new Locale("en", "UK"));
							cellvalue = simpledateformat.format(date);
//                 				    System.out.println("date cell value: "+cellvalue);
						}
						return cellvalue;
					}

					else if (cell.getCellType().toString().equals("BLANK"))

						return "";

					else {
						cellvalue = String.valueOf(cell.getBooleanCellValue());
						return cellvalue;
					}

				}

				catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return cellvalue;
	}

	public int rowcount(int sheetIndex) {

		/*
		 * THIS METHOD RETURNS THE VALUE AT A PARTICULAR CELL IN A SHEET.
		 * 
		 * Arguments - sheetIndex - takes integer value rowIndex - takes integer value
		 * colIndex -takes integer value
		 */
		int numberofrows;
		try {
			if (!(sheetIndex > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return 0;
			} else {
				// 1.
				/*
				 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
				 * will take the sheet number as integer value Index where 0 will indicate the
				 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
				 */
				sheet = wb.getSheetAt(sheetIndex - 1);

				if (sheet == null) {
					System.out.println("Sheet Doesn't Exist");
					return 0;
				}

				numberofrows = sheet.getLastRowNum();

			}
		}

		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		numberofrows = sheet.getLastRowNum();
		return (numberofrows + 1);

	}

	public int colcount(int sheetIndex) {

		/*
		 * THIS METHOD RETURNS THE VALUE AT A PARTICULAR CELL IN A SHEET.
		 * 
		 * Arguments - sheetIndex - takes integer value
		 */
		int numberofcolumns;
		try {
			if (!(sheetIndex > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return 0;
			} else {
				// 1.
				/*
				 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
				 * will take the sheet number as integer value Index where 0 will indicate the
				 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
				 */
				sheet = wb.getSheetAt(sheetIndex - 1);

				if (sheet == null) {
					System.out.println("Sheet Doesn't Exist");
					return 0;
				}

				row = sheet.getRow(0);
				if (row == null) {
					System.out.println("No column headers in the sheet");
					return 0;
				}

				numberofcolumns = row.getLastCellNum();

			}
		}

		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		numberofcolumns = row.getLastCellNum();
		return numberofcolumns;

	}

	public boolean insertNewrow(int sheetNumber, int rowindex) {

		/*
		 * THIS METHOD CREATES NEW ROW AT A PARTICULAR INDEX IN A SHEET.
		 * 
		 * Arguments - sheetNumber - 1. sheet in which row to be inserted 2. takes
		 * integer value rowIndex - 1. index number at which row to be inserted 2. takes
		 * integer value
		 */

		try {
			if (!(sheetNumber > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return false;
			} else {
				// 1.
				/*
				 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
				 * will take the sheet number as integer value Index where 0 will indicate the
				 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
				 */
				try {
					// 1.
					/*
					 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
					 * will take the sheet number as integer value Index where 0 will indicate the
					 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
					 */

					sheet = wb.getSheetAt(sheetNumber - 1);

					if (sheet == null) {
						System.out.println("Sheet Doesn't Exist");
						return false;
					}
					if (!((rowindex - 1) > 0)) {
						System.out.println("Cannot Insert data in row(" + rowindex + ")");
						return false;
					}

					// Insert a row at a particular index :
					row = sheet.createRow(rowindex - 1);

				}

				catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return true;
	}

	public boolean setcellData(int sheetNumber, int rowindex, int cellindex, String cellvalue) {
		/*
		 * THIS METHOD CREATES NEW ROW, CREATES A NEW CELL AND INSERT VALUE IN THAT CELL
		 * AT A PARTICULAR INDEX IN A SHEET ALONGWITH SOME CELL STYLINGS LIKE BACKGROUND
		 * COLOR,FOEGROUND COLOR etc.
		 * 
		 * Arguments - sheetNumber - 1. sheet in which row to be inserted 2. takes
		 * integer value rowIndex - 1. index number at which row to be inserted 2. takes
		 * integer value colIndex - 1. index number at which cell with data to be
		 * inserted 2. takes integer value
		 */

		try {
			if (!(sheetNumber > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return false;
			} else {
				// 1.
				/*
				 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
				 * will take the sheet number as integer value Index where 0 will indicate the
				 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
				 */
				try {
					sheet = wb.getSheetAt(sheetNumber - 1);

					if (sheet == null) {
						System.out.println("Sheet Doesn't Exist");
						return false;
					}
					if (!((rowindex - 1) > 0)) {
						System.out.println("Cannot Insert data in row(" + rowindex + ")");
						return false;
					} else if (!((cellindex - 1) >= 0)) {
						System.out.println("Cannot Insert data in cell(" + cellindex + ")");
						return false;
					} else {
						row = sheet.getRow(rowindex - 1);
						if (row == null) {
							// 2. Insert a row at a particular index :
							row = sheet.createRow(rowindex - 1);
						}
						// SET THE STYLES FOR THE CELLS ALONGWITH DIFFERENT COLORS:
//		                  style = wb.createCellStyle();
//		                  style.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());
//                          style.setFillPattern(FillPatternType.BIG_SPOTS);
//                          style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.index);

						// CREATE A CELL AT A PARTICULAR INDEX AND INSERT VALUE IN THE CELL :

						cell = row.createCell(cellindex - 1);
						cell.setCellValue(cellvalue);
						cell.setCellStyle(style);
					}

				} catch (Exception e) {

					e.printStackTrace();
				}

			}
			// Write out the value in the cells and close the file :

			out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		}

		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return true;
	}

	public boolean createcolumn(int sheetIndex, String colname) {
		/*
		 * THIS METHOD CREATES A NEW COLUMN WITH A PARTICULAR NAME Arguments -
		 * sheetINdex - 1. sheet in which column to be created 2. takes integer value
		 * colname - 1. name of the column to be created 2. takes String value
		 */
		try {
			if (!(sheetIndex > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return false;
			}
			// 1.
			/*
			 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
			 * will take the sheet number as integer value Index where 0 will indicate the
			 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
			 */
			else {
				try {
					sheet = wb.getSheetAt(sheetIndex - 1);

					row = sheet.getRow(0);

					if (row == null)

					{
//        			   System.out.println("-----BEFORE CREATING ROWS -----------------");
//        			   System.out.println("ROW VALUE(row): "+row);
//        			   System.out.println("First row in the sheet: "+sheet.getFirstRowNum());
//            		   System.out.println("Last row in the sheet: "+sheet.getLastRowNum());

						row = sheet.createRow(0);
					}
//        		   System.out.println("-----AFTER CREATING ROWS -----------------");
//        		   System.out.println("ROW VALUE(row): "+row);
//        		   System.out.println("First row in the sheet: "+sheet.getFirstRowNum());
//         		   System.out.println("Last row in the sheet: "+sheet.getLastRowNum());

					// System.out.println("ROW VALUE(String):
					// "+sheet.getRow(0).getCell(0).getStringCellValue());
					if (row.getLastCellNum() == -1) {
//        			   System.out.println("   -1   ");  
//                     System.out.println("-----BEFORE CREATING CELLS -----------------");
//        		     System.out.println("First Cell: "+row.getFirstCellNum());
//         		     System.out.println("Last Cell: "+row.getLastCellNum());
						cell = row.createCell(0);
					} else {
//                       System.out.println("  < -1   ");  
//                       System.out.println("-----BEFORE CREATING CELLS -----------------");
//                       System.out.println("First Cell: "+row.getFirstCellNum());
//        		       System.out.println("Last Cell: "+row.getLastCellNum());
						cell = row.createCell(row.getLastCellNum());
					}
//        		   System.out.println("-----AFTER CREATING CELLS --------------------");
//        		   System.out.println("First Cell: "+row.getFirstCellNum());
//         		   System.out.println("Last Cell: "+row.getLastCellNum());
					cell.setCellValue(colname);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			out = new FileOutputStream(file);
			wb.write(out);
			out.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

		return true;
	}

	public boolean removecolumn(int sheetIndex, int colIndex) {
		/*
		 * THIS METHOD REMOVES A COLUMN WITH A PARTICULAR INDEX Arguments - sheetINdex -
		 * 1. sheet in which column to be created 2. takes integer value colname - 1.
		 * name of the column to be created 2. takes String value
		 */
		try {
			if (!(sheetIndex > 0)) {
				System.out.println("Invalid sheet Index-Please Try again");
				return false;
			}
			// 1.
			/*
			 * a) Load a specific sheet in the workbook : b) getsheetAt(int index) method
			 * will take the sheet number as integer value Index where 0 will indicate the
			 * sheet number 1 in workbook, 1 indicates sheet number 2 and so on.
			 */
			else {
				try {
					sheet = wb.getSheetAt(sheetIndex - 1);
					row = sheet.getRow(0);
					if (sheet.getLastRowNum() == -1) {
//        			   System.out.println("----------IN IF -------------");
//        			   System.out.println("ROW: "+row);
//       			   System.out.println("Last row: "+sheet.getLastRowNum());       			       
						System.out.println("Row Doesn't exist or Blank row");
						return false;
					}

					else if (!((colIndex - 1) >= 0)) {
						System.out.println("Cannot remove data from cell(" + colIndex + ")");
						return false;
					} else {
//        			  System.out.println("----------IN ELSE -------------");
//        			  System.out.println("ROW: "+row);
//   			          System.out.println("Last row: "+sheet.getLastRowNum());
//   			       System.out.println("---------------IN LOOP ------------------");
						for (int i = 0; i < (sheet.getLastRowNum() + 1); i++) {

							row = sheet.getRow(i);

							if (row == null) {
								row = sheet.getRow(i);
							} else if (row.getCell(colIndex - 1) == null) {
								row = sheet.getRow(i);
							} else {
								row.removeCell(row.getCell(colIndex - 1));
							}
						}
					}

				} catch (Exception e) {
//		              System.out.println("Sheet Doesn't exist");
					e.printStackTrace();
				}
			}

			out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return true;
	}
}
