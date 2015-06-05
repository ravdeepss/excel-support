package com.excelsupport.writer;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.excelsupport.ExcelWorkbookWrapper;

public interface Writer
{

	void write (ExcelWorkbookWrapper swb) throws FileNotFoundException, IOException;

	void writeRowMap (ExcelWorkbookWrapper swb) throws IOException;
}
