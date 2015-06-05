package com.excelsupport;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 * 
 */
public class ExcelWorkbookWrapper
{
	private ExcelRow				headerRow;
	private Map<Integer, ExcelRow>	rowMap;
	private String					sheetName;

	public void setHeaderRow (ExcelRow headerRow2)
	{
		this.headerRow = headerRow2;

	}

	public void setRowMap (Map<Integer, ExcelRow> rowMap2)
	{
		this.rowMap = rowMap2;

	}

	public Map<Integer, ExcelRow> getRowsInRange (int start, int end)
	{

		Map<Integer, ExcelRow> retRowMap = new HashMap<Integer, ExcelRow> ();

		if (this.rowMap.containsKey (start) && this.rowMap.containsKey (end))
		{
			for (int i = start; i <= end; i++)
			{
				ExcelRow sr = this.rowMap.get (i);
				if (sr != null)
				{
					retRowMap.put (i, sr);
				}

			}
		}
		return retRowMap;

	}

	public String getCellValue (int rowNum, String columnName)
	{
		String retVal = "Not Found";

		retVal = rowMap.get (rowNum).getCellValue (columnName.toUpperCase ());

		return retVal;
	}

	@Override
	public String toString ()
	{

		StringBuffer sb = new StringBuffer ();
		sb.append ("HEADER|" + headerRow.toString ()).append ("\n");
		for (ExcelRow sr : rowMap.values ())
		{
			sb.append (sr.getRowNum () + "|" + sr.toString ()).append ("\n");
		}

		return sb.toString ();
	}

	/**
	 * 
	 * @return The Header Row of the Parsed Spreadsheet
	 */
	public ExcelRow getHeaderRow ()
	{
		return headerRow;
	}

	public int getNumberOfRowsExcludingHeader ()
	{
		return rowMap.size ();
	}

	public Collection<ExcelRow> getAllRowsExcludingHeader(){
		return rowMap.values();
	}
	
	public void setSheetName (String sheetName)
	{
		this.sheetName = sheetName;

	}

	public String getSheetName ()
	{
		return sheetName;
	}

}
