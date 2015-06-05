package com.excelsupport.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.excelsupport.ExcelRow;
import com.excelsupport.ExcelWorkbookWrapper;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 */
public class ExcelReader implements Reader
{
	private Logger		debug	= Logger.getLogger (ExcelReader.class);
	private File		file;
	private InputStream	is;
	private String		rowName	= "";

	
	/**
	 * Default constructor. Use the read(String filePath) method if using this constructor
	 */
	public ExcelReader(){
		
	}
	
	/**
	 * Initialize the reader with an Input Stream
	 * @param is
	 */
	public ExcelReader (InputStream is)
	{
		this.is = is;
	}
	/**
	 * Initialize the reader with a File object
	 * @param f
	 */
	public ExcelReader (File f)
	{
		this.file = f;
	}


	public ExcelWorkbookWrapper read (String filePath)
	{
		this.file = new File (filePath);
		return read ();
	}

	public ExcelWorkbookWrapper read ()
	{
		if (is == null)
		{
			try
			{
				is = new FileInputStream (file);

			}
			catch (FileNotFoundException e)
			{
				debug.error ("File not found in the specified path.");
				debug.error (e);
			}
		}
		ExcelWorkbookWrapper wb = new ExcelWorkbookWrapper ();
		POIFSFileSystem fileSystem = null;

		try
		{
			fileSystem = new POIFSFileSystem (is);

			HSSFWorkbook workBook = new HSSFWorkbook (fileSystem);
			HSSFSheet sheet = workBook.getSheetAt (0);

			wb.setSheetName (sheet.getSheetName ());
			Iterator<Row> rows = sheet.rowIterator ();

			ExcelRow headerRow = createHeaderRow (sheet);

			wb.setHeaderRow (headerRow);
			int lastRow = sheet.getLastRowNum ();
			Map<Integer, ExcelRow> rowMap = new HashMap<Integer, ExcelRow> ();

			for (int i = 1; i <= lastRow; i++)
			{
				if(sheet.getRow (i)!=null)
				{	
					HSSFRow row = sheet.getRow (i);

					debug.info ("Row No.: " + row.getRowNum ());

					rowMap.put (row.getRowNum (), processRow (row, headerRow));
				}
			}
			wb.setRowMap (rowMap);

		}
		catch (Exception e)
		{
			debug.error ( e.getMessage (),e);
		}
		return wb;
	}

	private ExcelRow processRow (HSSFRow row, ExcelRow headerRow)
	{

		// once get a row its time to iterate through cells.
		Iterator<Cell> cells = row.cellIterator ();
		ExcelRow sr = new ExcelRow ();
		sr.setRowNum (row.getRowNum ());
		sr.setRowName (rowName);
		Map<String, String> cellMap = new HashMap<String, String> ();

		while (cells.hasNext ())
		{
			HSSFCell cell = (HSSFCell) cells.next ();
			cell.setCellType (HSSFCell.CELL_TYPE_STRING);
			Integer colNo = cell.getColumnIndex ();
			String headerName = headerRow.getCellValue (colNo.toString ());
			String cellValue = "";

			/*
			 * Now we will get the cell type and display the values accordingly.
			 */
			switch (cell.getCellType ())
			{
				case HSSFCell.CELL_TYPE_NUMERIC:
				{

					// cell type numeric.
					debug.info ("Numeric value: " + cell.getNumericCellValue ());

					cellValue = cell.getNumericCellValue () + "";
					break;
				}

				case HSSFCell.CELL_TYPE_STRING:
				{

					// cell type string.
					HSSFRichTextString richTextString = cell.getRichStringCellValue ();

					debug.info ("String value: " + richTextString.getString ());
					cellValue = richTextString.getString ();
					break;
				}

				default:
				{

					// types other than String and Numeric.
					debug.info ("Type not supported.");
					cellValue = "Not Supported";
					break;
				}
			}
			debug.info ("Cell No.: " + headerName + ", Cell Value: " + cellValue);
			cellMap.put (headerName, cellValue.trim ());

		}
		sr.setCells (cellMap);
		return sr;
	}

	private ExcelRow createHeaderRow (HSSFSheet sheet) throws Exception
	{
		HSSFRow row = sheet.getRow (0);

		Iterator<Cell> cells = row.cellIterator ();
		ExcelRow sr = new ExcelRow ();
		sr.setRowNum (row.getRowNum ());
		Map<String, String> cellMap = new LinkedHashMap<String, String> ();

		while (cells.hasNext ())
		{
			HSSFCell cell = (HSSFCell) cells.next ();

			System.out.println ("Cell No.: " + cell.getColumnIndex () + ", Cell Value: " + cell.getStringCellValue ());
			Integer colNo = cell.getColumnIndex ();
			String cellValue = cell.getStringCellValue ().trim ();
			if (StringUtils.isBlank (cellValue))
			{
				throw new Exception ("Header cannot be blank: " + colNo);
			}
			cellMap.put (colNo.toString (), cellValue.toUpperCase ());

		}
		sr.setCells (cellMap);
		sr.setRowName ("HEADER");
		return sr;
	}

	public File getFile ()
	{
		return file;
	}

	public void setFile (File file)
	{
		this.file = file;
	}

	public String getRowName ()
	{
		return rowName;
	}

	public void setRowName (String rowName)
	{
		this.rowName = rowName;
	}

}
