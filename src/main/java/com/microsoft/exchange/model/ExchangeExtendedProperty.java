/**
 * 
 */
package com.microsoft.exchange.model;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.MapiPropertyTypeType;
import com.microsoft.exchange.types.PathToExtendedFieldType;

/**
 * @author ctcudd
 *
 */
public interface ExchangeExtendedProperty<V> {

	public String getICalPropertyName();
	
	public String getExchangePropertyName();
	
	public String getExchangePropertySetId();
	
	public MapiPropertyTypeType getExchangePropertyType();
	
	public PathToExtendedFieldType getPathToExtendedFieldType();
	
	public ExtendedPropertyType constructExtendedProperty(V value);
	
	public V extractPropertyValue(ItemType item);
	
}
