package com.ravz.playground.excelsupport;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 * 
 */
public class SimpleWorkBook
{
	private SimpleRow				headerRow;
	private Map<Integer, SimpleRow>	rowMap;
	private String					sheetName;

	public void setHeaderRow (SimpleRow headerRow2)
	{
		this.headerRow = headerRow2;

	}

	public void setRowMap (Map<Integer, SimpleRow> rowMap2)
	{
		this.rowMap = rowMap2;

	}

	public Map<Integer, SimpleRow> getRowsInRange (int start, int end)
	{

		Map<Integer, SimpleRow> retRowMap = new HashMap<Integer, SimpleRow> ();

		if (this.rowMap.containsKey (start) && this.rowMap.containsKey (end))
		{
			for (int i = start; i <= end; i++)
			{
				SimpleRow sr = this.rowMap.get (i);
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
		for (SimpleRow sr : rowMap.values ())
		{
			sb.append (sr.getRowNum () + "|" + sr.toString ()).append ("\n");
		}

		return sb.toString ();
	}

	/**
	 * 
	 * @return The Header Row of the Parsed Spreadsheet
	 */
	public SimpleRow getHeaderRow ()
	{
		return headerRow;
	}

	public int getNumberOfRowsExcludingHeader ()
	{
		return rowMap.size ();
	}

	public Collection<SimpleRow> getAllRowsExcludingHeader(){
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
