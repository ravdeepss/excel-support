package com.excelsupport;

import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 */
public class SupportUtils
{

	/**
	 * Utility method to help split a large workbook into a smaller number of workbooks
	 * 
	 * @param wb
	 *            SimpleWorkBook Instance
	 * @param numberOfBatches
	 *            The number of batches this workbook should be split in
	 * @return
	 */
	public static List<ExcelWorkbookWrapper> splitIntoBatches (ExcelWorkbookWrapper wb, int numberOfBatches)
	{
		List<ExcelWorkbookWrapper> books = new ArrayList<ExcelWorkbookWrapper> ();
		int totalRows = wb.getNumberOfRowsExcludingHeader();// 56
		int batchSize = totalRows / numberOfBatches; // 56/10= 5
		int lastBatchSize = totalRows % numberOfBatches;// 6
		if (batchSize <= 1)
		{
			books.add (wb);
		}
		else
		{
			ExcelRow header = wb.getHeaderRow ();
			int start = 2;
			int end = (start + batchSize) - 1;
			for (int i = 1; i < numberOfBatches; i++)
			{
				Map<Integer, ExcelRow> rowMap = wb.getRowsInRange (start, end); // 2,11
				ExcelWorkbookWrapper book = new ExcelWorkbookWrapper ();
				book.setHeaderRow (header);
				book.setSheetName (wb.getSheetName ());
				book.setRowMap (rowMap);
				books.add (book);
				start = end + 1;
				end = (start + batchSize) - 1;
			}
			if (lastBatchSize != 0)
			{
				end = (start + lastBatchSize) - 1;
				Map<Integer, ExcelRow> rowMap = wb.getRowsInRange (start, end);
				ExcelWorkbookWrapper book = new ExcelWorkbookWrapper ();
				book.setHeaderRow (header);
				book.setSheetName (wb.getSheetName ());
				book.setRowMap (rowMap);
				books.add (book);
			}
		}
		return books;
	}

	public static <T> Map<Integer, ExcelRow> buildRowMap (Collection<T> objs, RowTransformer<T> t, ExcelRow headerRow)
	{
		Map<Integer, ExcelRow> rowMap = new LinkedHashMap<Integer, ExcelRow> ();
		Integer i = 1;
		for (T obj : objs)
		{
			ExcelRow sr = t.transform (obj, headerRow);
			sr.setRowNum (i);
			rowMap.put (i++, sr);
		}
		return rowMap;
	}
}
