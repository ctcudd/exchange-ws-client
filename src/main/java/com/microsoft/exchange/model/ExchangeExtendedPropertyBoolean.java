/**
 * 
 */
package com.microsoft.exchange.model;

import org.apache.commons.lang.BooleanUtils;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.MapiPropertyTypeType;

/**
 * @author ctcudd
 *
 */
public abstract class ExchangeExtendedPropertyBoolean extends AbstractExchangeExtendedProperty<Boolean> {

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#getExchangePropertyType()
	 */
	@Override
	public final MapiPropertyTypeType getExchangePropertyType() {
		return MapiPropertyTypeType.BOOLEAN;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#constructExtendedProperty(java.lang.Object)
	 */
	@Override
	public ExtendedPropertyType constructExtendedProperty(Boolean value) {
		ExtendedPropertyType extendedProperty = new ExtendedPropertyType();
		extendedProperty.setExtendedFieldURI(this.getPathToExtendedFieldType());
		extendedProperty.setValue(BooleanUtils.toStringTrueFalse(value));
		return extendedProperty;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#extractPropertyValue(com.microsoft.exchange.types.ItemType)
	 */
	@Override
	public Boolean extractPropertyValue(ItemType item) {
		Boolean extendedPropertyValue = null;
		ExtendedPropertyType extendedProperty = this.extractExtendedPropertyType(item);
		if(null != extendedProperty){
			extendedPropertyValue = BooleanUtils.toBooleanObject(extendedProperty.getValue());
		}
		return extendedPropertyValue;
	}

}
