/**
 * 
 */
package com.microsoft.exchange.model;

import java.util.Collection;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.MapiPropertyTypeType;
import com.microsoft.exchange.types.NonEmptyArrayOfPropertyValuesType;

/**
 * @author ctcudd
 *
 */
public abstract class ExchangeExtendedPropertyStringArray extends AbstractExchangeExtendedProperty<Collection<String>> {


	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#getExchangePropertyType()
	 */
	@Override
	public final MapiPropertyTypeType getExchangePropertyType() {
		return MapiPropertyTypeType.STRING_ARRAY;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#constructExtendedProperty(java.lang.Object)
	 */
	@Override
	public ExtendedPropertyType constructExtendedProperty(Collection<String> values) {
		ExtendedPropertyType extendedProperty = new ExtendedPropertyType();
		extendedProperty.setExtendedFieldURI(this.getPathToExtendedFieldType());
		NonEmptyArrayOfPropertyValuesType valuesArray = new NonEmptyArrayOfPropertyValuesType();
		valuesArray.getValues().addAll(values);
		extendedProperty.setValues(valuesArray);
		return extendedProperty;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#extractPropertyValue(com.microsoft.exchange.types.ItemType)
	 */
	@Override
	public Collection<String> extractPropertyValue(ItemType item) {
		Collection<String> extendedPropertyValue = null;
		ExtendedPropertyType extendedProperty = this.extractExtendedPropertyType(item);
		if(null != extendedProperty){
			extendedPropertyValue = extendedProperty.getValues().getValues();
		}
		return extendedPropertyValue;
	}

}
