package com.ravz.playground.excelsupport.reader;

import com.ravz.playground.excelsupport.SimpleWorkBook;

/**
 * 
 * @author Ravdeep Sandhu
 * 
 */
public interface Reader
{

	/**
	 * Returns a {@link SimpleWorkBook} 
	 * The Reader is expected to be initialized using the constructor if this method is to be used
	 * @param file
	 * @return {@link SimpleWorkBook} 
	 */
	public SimpleWorkBook read ();

	/**
	 * This method will initilize the file object using the filepath provided and return a {@link SimpleWorkBook} object
	 * @param filePath
	 * @return {@link SimpleWorkBook}
	 */
	public SimpleWorkBook read (String filePath);

}
