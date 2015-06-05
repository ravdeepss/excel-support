package com.excelsupport;

public interface SimpleRowTransformer<T>
{

	public SimpleRow transform (T dto, SimpleRow headerRow);

}
