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
public abstract class ExchangeExtendedPropertyString extends AbstractExchangeExtendedProperty<String> {


	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#getExchangePropertyType()
	 */
	@Override
	public final MapiPropertyTypeType getExchangePropertyType() {
		return MapiPropertyTypeType.STRING;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#constructExtendedProperty(java.lang.Object)
	 */
	@Override
	public ExtendedPropertyType constructExtendedProperty(String value) {
		ExtendedPropertyType extendedProperty = new ExtendedPropertyType();
		extendedProperty.setExtendedFieldURI(this.getPathToExtendedFieldType());
		extendedProperty.setValue(value);
		return extendedProperty;
	}

	@Override
	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#extractPropertyValue(com.microsoft.exchange.types.CalendarItemType)
	 */
	public String extractPropertyValue(ItemType item) {
		String extendedPropertyValue = null;
		ExtendedPropertyType extendedProperty = this.extractExtendedPropertyType(item);
		if(null != extendedProperty){
			extendedPropertyValue = extendedProperty.getValue();
		}
		return extendedPropertyValue;
	}
}
