package com.ravz.playground.excelsupport;

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
public class SimpleWorkBookHelper
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
	public static List<SimpleWorkBook> splitIntoBatches (SimpleWorkBook wb, int numberOfBatches)
	{
		List<SimpleWorkBook> books = new ArrayList<SimpleWorkBook> ();
		int totalRows = wb.getNumberOfRowsExcludingHeader();// 56
		int batchSize = totalRows / numberOfBatches; // 56/10= 5
		int lastBatchSize = totalRows % numberOfBatches;// 6
		if (batchSize <= 1)
		{
			books.add (wb);
		}
		else
		{
			SimpleRow header = wb.getHeaderRow ();
			int start = 2;
			int end = (start + batchSize) - 1;
			for (int i = 1; i < numberOfBatches; i++)
			{
				Map<Integer, SimpleRow> rowMap = wb.getRowsInRange (start, end); // 2,11
				SimpleWorkBook book = new SimpleWorkBook ();
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
				Map<Integer, SimpleRow> rowMap = wb.getRowsInRange (start, end);
				SimpleWorkBook book = new SimpleWorkBook ();
				book.setHeaderRow (header);
				book.setSheetName (wb.getSheetName ());
				book.setRowMap (rowMap);
				books.add (book);
			}
		}
		return books;
	}

	public static <T> Map<Integer, SimpleRow> buildRowMap (Collection<T> objs, SimpleRowTransformer<T> t, SimpleRow headerRow)
	{
		Map<Integer, SimpleRow> rowMap = new LinkedHashMap<Integer, SimpleRow> ();
		Integer i = 1;
		for (T obj : objs)
		{
			SimpleRow sr = t.transform (obj, headerRow);
			sr.setRowNum (i);
			rowMap.put (i++, sr);
		}
		return rowMap;
	}
}
