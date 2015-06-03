package com.ravz.playground.excelsupport.writer;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.ravz.playground.excelsupport.SimpleWorkBook;

public interface Writer
{

	void write (SimpleWorkBook swb) throws FileNotFoundException, IOException;

	void writeRowMap (SimpleWorkBook swb) throws IOException;
}
