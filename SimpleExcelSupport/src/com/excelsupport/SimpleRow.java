package com.excelsupport;

import java.util.Collection;
import java.util.Map;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 */
public class SimpleRow
{
	private Integer				rowNum;
	private Map<String, String>	cells;
	private String				rowName;

	public Integer getRowNum ()
	{
		return rowNum;
	}

	public void setRowNum (Integer rowNum)
	{
		this.rowNum = rowNum;
	}

	public Map<String, String> getCells ()
	{
		return cells;
	}

	public void setCells (Map<String, String> cells)
	{
		this.cells = cells;
	}

	public String getCellValue (String columnName)
	{

		return cells.get (columnName.toUpperCase ());
	}

	@Override
	public String toString ()
	{
		Collection<String> values = cells.values ();
		StringBuffer sb = new StringBuffer ();

		for (String val : values)
		{
			sb.append (val).append ("|");
		}

		return sb.toString ();
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
