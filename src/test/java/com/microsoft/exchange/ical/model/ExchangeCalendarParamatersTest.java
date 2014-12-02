/**
 * 
 */
package com.microsoft.exchange.ical.model;

import static junit.framework.Assert.*;
import net.fortuna.ical4j.model.parameter.XParameter;

import org.junit.Test;

import com.microsoft.exchange.types.MailboxTypeType;
import com.microsoft.exchange.types.MapiPropertyTypeType;
import com.microsoft.exchange.types.PathToExtendedFieldType;

/**
 * @author ctcudd
 *
 */
public class ExchangeCalendarParamatersTest {

	/**
	 * test to verify all {@link MailboxTypeType} produce a valid {@link XParameter}
	 */
	@Test
	public void emailAddressMailboxTypeParamater_control(){
		for(MailboxTypeType type : MailboxTypeType.values()){
			EmailAddressMailboxTypeParamater param = new EmailAddressMailboxTypeParamater(type);
			assertNotNull(param);
			assertEquals(EmailAddressMailboxTypeParamater.EMAIL_ADDRESS_MAILBOX_TYPE, param.getName());
			assertNotNull(param.getValue());
			assertEquals(type.value(), param.getValue());
		}
	}
	
	/**
	 * Verify a routingType string produces a valid {@link EmailAddressRoutingTypeParamater}
	 */
	@Test
	public void emailAddressRoutingTypeParamater_control(){
		EmailAddressRoutingTypeParamater param = new EmailAddressRoutingTypeParamater(null);
		assertNotNull(param);
		assertEquals(EmailAddressRoutingTypeParamater.EMAIL_ADDRESS_ROUTING_TYPE, param.getName());
		assertNull(param.getValue());
		
		param = new EmailAddressRoutingTypeParamater("SMTP");
		assertNotNull(param);
		assertEquals(EmailAddressRoutingTypeParamater.EMAIL_ADDRESS_ROUTING_TYPE, param.getName());
		assertEquals("SMTP", param.getValue());
	}
	
	/**
	 * Verify params can be obtained from {@link PathToExtendedFieldType}
	 */
	@Test
	public void extendedFieldTypeParamaters_control(){
		int propertyId = 123456789;
		String propertySetId = "qwertyuiop";
		String propertyTag = "asdfghjkl";
		MapiPropertyTypeType propertyType = MapiPropertyTypeType.CLSID_ARRAY;
		
		PathToExtendedFieldType path = new PathToExtendedFieldType();
		path.setPropertyId(propertyId);
		path.setPropertySetId(propertySetId);
		path.setPropertyTag(propertyTag);
		path.setPropertyType(propertyType);
		
		PathToExtendedFieldTypePropertyIdParamater propertyIdParam = new PathToExtendedFieldTypePropertyIdParamater(path);
		assertNotNull(propertyIdParam);
		assertEquals(PathToExtendedFieldTypePropertyIdParamater.PATH_TO_EXTENDED_FIELD_TYPE_PROPERTY_ID, propertyIdParam.getName());
		assertEquals(propertyId, new Integer(propertyIdParam.getValue()).intValue());
		
		PathToExtendedFieldTypePropertySetIdParamater propertySetIdParam = new PathToExtendedFieldTypePropertySetIdParamater(path);
		assertNotNull(propertySetIdParam);
		assertEquals(PathToExtendedFieldTypePropertySetIdParamater.PATH_TO_EXTENDED_FIELD_TYPE_PROPERTY_SET_ID, propertySetIdParam.getName());
		assertEquals(propertySetId, propertySetIdParam.getValue());
		
		PathToExtendedFieldTypePropertyTagParamater propertyTagParam = new PathToExtendedFieldTypePropertyTagParamater(path);
		assertNotNull(propertyTagParam);
		assertEquals(PathToExtendedFieldTypePropertyTagParamater.PATH_TO_EXTENDED_FIELD_TYPE_PROPERTY_TAG, propertyTagParam.getName());
		assertEquals(propertyTag, propertyTagParam.getValue());
		
		PathToExtendedFieldTypePropertyTypeParamater propertyTypeParam = new PathToExtendedFieldTypePropertyTypeParamater(path);
		assertNotNull(propertyTypeParam);
		assertEquals(PathToExtendedFieldTypePropertyTypeParamater.PATH_TO_EXTENDED_FIELD_TYPE_PROPERTY_TYPE, propertyTypeParam.getName());
		assertEquals(propertyType.value(), propertyTypeParam.getValue());
	}
	
}
