package com.excelsupport.reader;

import com.excelsupport.ExcelWorkbookWrapper;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 */
public interface Reader
{

	/**
	 * Returns a {@link ExcelWorkbookWrapper} 
	 * The Reader is expected to be initialized using the constructor if this method is to be used
	 * @param file
	 * @return {@link ExcelWorkbookWrapper} 
	 */
	public ExcelWorkbookWrapper read ();

	/**
	 * This method will initilize the file object using the filepath provided and return a {@link ExcelWorkbookWrapper} object
	 * @param filePath
	 * @return {@link ExcelWorkbookWrapper}
	 */
	public ExcelWorkbookWrapper read (String filePath);

}
