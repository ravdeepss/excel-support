package com.excelsupport.writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.excelsupport.SimpleRow;
import com.excelsupport.SimpleWorkBook;

public class ExcelWriter implements Writer
{
	private Logger			debug	= Logger.getLogger (ExcelWriter.class);
	private File			targetFile;
	private InputStream		is;
	private OutputStream	os;

	public ExcelWriter (File f)
	{
		this.targetFile = f;
	}

	@Override
	public void write (SimpleWorkBook swb) throws IOException
	{

		HSSFWorkbook workBook = writeAllRowsExcludingHeader (swb);
		createHeaderRow (workBook.getSheetAt (0), swb.getHeaderRow ());
		if (os == null)
		{
			os = new FileOutputStream (targetFile);
		}
		workBook.write (os);
	}

	/**
	 * @param workBook
	 */
	private HSSFWorkbook writeAllRowsExcludingHeader (SimpleWorkBook swb)
	{
		POIFSFileSystem fileSystem = null;
		HSSFWorkbook workBook = null;
		try
		{
			fileSystem = getPoiFileSystem ();
			workBook = new HSSFWorkbook (fileSystem);
			HSSFSheet sheet = workBook.getSheetAt (0);

			Map<Integer, SimpleRow> rowMap = swb.getRowsInRange (1, swb.getNumberOfRowsExcludingHeader ());
			for (Entry<Integer, SimpleRow> entry : rowMap.entrySet ())
			{
				HSSFRow row = sheet.createRow (entry.getKey ());
				debug.info ("Row No.: " + row.getRowNum ());
				processRow (row, entry.getValue (), swb.getHeaderRow ());
			}

		}
		catch (Exception e)
		{
			debug.error ( e.getMessage (),e);
		}

		return workBook;
	}

	private POIFSFileSystem getPoiFileSystem () throws IOException
	{
		if (is == null)
		{
			try
			{
				is = new FileInputStream (targetFile);

			}
			catch (FileNotFoundException e)
			{
				debug.error ("File not found in the specified path.");
				debug.error (e);
			}
		}
		return new POIFSFileSystem (is);
	}

	private void createHeaderRow (HSSFSheet sheet, SimpleRow headerRow)
	{
		HSSFRow row = sheet.createRow (0);
		for (Entry<String, String> entry : headerRow.getCells ().entrySet ())
		{
			HSSFCell cell = row.createCell (Integer.parseInt (entry.getKey ()));
			cell.setCellValue (entry.getValue ());
		}
	}

	@Override
	public void writeRowMap (SimpleWorkBook swb) throws IOException
	{
		HSSFWorkbook workBook = writeAllRowsExcludingHeader (swb);
		if (os == null)
		{
			os = new FileOutputStream (targetFile);
		}
		workBook.write (os);
	}

	private void processRow (HSSFRow row, SimpleRow rowToWrite, SimpleRow headerRow)
	{

		for (Entry<String, String> entry : headerRow.getCells ().entrySet ())
		{
			HSSFCell cell = row.createCell (Integer.parseInt (entry.getKey ()));
			cell.setCellValue (rowToWrite.getCellValue (entry.getValue ()));
			debug.info ("Cell No.: " + entry.getValue () + ", Cell Value: " + rowToWrite.getCellValue (entry.getValue ()));
		}
	}

}
