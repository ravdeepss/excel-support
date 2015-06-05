package com.excelsupport;

public interface RowTransformer<T>
{

	public ExcelRow transform (T dto, ExcelRow headerRow);

}
