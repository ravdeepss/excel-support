package com.excelsupport.writer;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.excelsupport.SimpleWorkBook;

public interface Writer
{

	void write (SimpleWorkBook swb) throws FileNotFoundException, IOException;

	void writeRowMap (SimpleWorkBook swb) throws IOException;
}
