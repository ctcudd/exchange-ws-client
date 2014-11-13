/**
 * 
 */
package com.microsoft.exchange.model;

import java.util.List;

import org.apache.commons.lang.BooleanUtils;
import org.springframework.util.CollectionUtils;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.PathToExtendedFieldType;


/**
 * @author ctcudd
 *
 */
public abstract class AbstractExchangeExtendedProperty<V> implements ExchangeExtendedProperty<V> {

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.exchange.model.ExchangeExtendedProperty#getPathToExtendedFieldType()
	 */
	@Override
	public PathToExtendedFieldType getPathToExtendedFieldType() {
		PathToExtendedFieldType extendedFieldType = new PathToExtendedFieldType();
		extendedFieldType.setPropertyName(this.getExchangePropertyName());
		extendedFieldType.setPropertySetId(this.getExchangePropertySetId());
		extendedFieldType.setPropertyType(this.getExchangePropertyType());
		return extendedFieldType;
	}
	
	/**
	 * Given an {@link ItemType}, loop through {@link ExtendedPropertyType}s and return the first that matches.
	 * A matching {@link ExtendedPropertyType} has the same propertyName, propertySetId, and propertyType.
	 * @param item
	 * @return the matching {@link ExtendedPropertyType} if found, otherwise <code>null</code>.
	 */
	protected ExtendedPropertyType extractExtendedPropertyType(ItemType item){
		ExtendedPropertyType prop = null;
		if(null != item){
			List<ExtendedPropertyType> extendedProperties = item.getExtendedProperties();
			if(!CollectionUtils.isEmpty(extendedProperties)){
				for(ExtendedPropertyType exProp: extendedProperties){
					if(null != exProp && null != exProp.getExtendedFieldURI()){
						PathToExtendedFieldType extendedFieldURI = exProp.getExtendedFieldURI();
						boolean nameMatch  = extendedFieldURI.getPropertyName().equals(this.getExchangePropertyName());
						boolean setIdMatch = extendedFieldURI.getPropertySetId().equals(this.getExchangePropertySetId());
						boolean typeMatch  = extendedFieldURI.getPropertyType().equals(this.getExchangePropertyType());
						
						if(nameMatch && setIdMatch && typeMatch){
							prop = exProp;
						}
					}
				}
			}
		}
		return prop;
	}
}
