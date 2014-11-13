/**
 * 
 */
package com.microsoft.exchange.model;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.MapiPropertyTypeType;

/**
 * @author ctcudd
 *
 */
public abstract class ExchangeExtendedPropertyInteger extends AbstractExchangeExtendedProperty<Integer> {

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#getExchangePropertyType()
	 */
	@Override
	public final MapiPropertyTypeType getExchangePropertyType() {
		return MapiPropertyTypeType.INTEGER;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#constructExtendedProperty(java.lang.Object)
	 */
	@Override
	public ExtendedPropertyType constructExtendedProperty(Integer value) {
		ExtendedPropertyType extendedProperty = new ExtendedPropertyType();
		extendedProperty.setExtendedFieldURI(this.getPathToExtendedFieldType());
		extendedProperty.setValue(value.toString());
		return extendedProperty;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#extractPropertyValue(com.microsoft.exchange.types.ItemType)
	 */
	@Override
	public Integer extractPropertyValue(ItemType item) {
		Integer value = null;
		ExtendedPropertyType extendedProperty = this.extractExtendedPropertyType(item);
		if(null != extendedProperty){
			value = Integer.parseInt(extendedProperty.getValue());
		}
		return value;
	}

}
