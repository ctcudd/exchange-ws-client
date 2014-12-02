/**
 * 
 */
package com.microsoft.exchange.ical.model;

import static junit.framework.Assert.assertEquals;
import static junit.framework.Assert.assertNotNull;

import org.junit.Test;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.PathToExtendedFieldType;
/**
 * @author ctcudd
 *
 */
public class ExchangeCalendarPropertiesTest {

	/**
	 * Verify {@link ExchangeTimeZoneProperty} can be obtained from a string
	 * representation of a timezone id
	 */
	@Test
	public void exchangeTimeZoneProperties_control(){
		String tzid = "cst";
		
		ExchangeTimeZoneProperty tzProp = new ExchangeTimeZoneProperty(tzid);
		assertNotNull(tzProp);
		assertEquals(ExchangeTimeZoneProperty.X_EWS_TIMEZONE, tzProp.getName());
		assertEquals(tzid, tzProp.getValue());
		
		ExchangeEndTimeZoneProperty endTzProp = new ExchangeEndTimeZoneProperty(tzid);
		assertNotNull(endTzProp);
		assertEquals(ExchangeEndTimeZoneProperty.X_EWS_END_TIMEZONE, endTzProp.getName());
		assertEquals(tzid, endTzProp.getValue());
		
		ExchangeStartTimeZoneProperty startTzProp = new ExchangeStartTimeZoneProperty(tzid);
		assertNotNull(startTzProp);
		assertEquals(ExchangeStartTimeZoneProperty.X_EWS_START_TIMEZONE, startTzProp.getName());
		assertEquals(tzid, startTzProp.getValue());
	}
	
	/**
	 * Verify an {@link ItemTypeItemIdProperty} and an
	 * {@link ItemTypeChangeKeyProperty} can be obtained from an
	 * {@link ItemIdType}
	 */
	@Test
	public void exchangeItemIdProperties_control(){
		ItemIdType itemId =  new ItemIdType();
		itemId.setId("1234567890");
		itemId.setChangeKey("asdfghjkl");
		
		ItemTypeItemIdProperty itemIdProp = new ItemTypeItemIdProperty(itemId);
		assertNotNull(itemIdProp);
		assertEquals(ItemTypeItemIdProperty.X_EWS_ITEM_ID, itemIdProp.getName());
		assertEquals(itemId.getId(), itemIdProp.getValue());
		
		ItemTypeChangeKeyProperty changeKeyProp = new ItemTypeChangeKeyProperty(itemId);
		assertNotNull(changeKeyProp);
		assertEquals(ItemTypeChangeKeyProperty.X_EWS_ITEM_CHANGEKEY, changeKeyProp.getName());
		assertEquals(itemId.getChangeKey(), changeKeyProp.getValue());
	}
	
	@Test
	public void exchangeItemParentFolderIdProperties_control(){
		FolderIdType fid = new FolderIdType();
		fid.setId("1234567890");
		fid.setChangeKey("asdfghjkl");
		
		ItemTypeParentFolderIdProperty parentFidProp = new ItemTypeParentFolderIdProperty(fid);
		assertNotNull(parentFidProp);
		assertEquals(ItemTypeParentFolderIdProperty.X_EWS_PARENT_FOLDER_ID, parentFidProp.getName());
		assertEquals(fid.getId(), parentFidProp.getValue());
		
		ItemTypeParentFolderChangeKeyProperty parentFkeyProp = new ItemTypeParentFolderChangeKeyProperty(fid);
		assertNotNull(parentFkeyProp);
		assertEquals(ItemTypeParentFolderChangeKeyProperty.X_EWS_PARENT_FOLDER_CHANGEKEY, parentFkeyProp.getName());
		assertEquals(fid.getChangeKey(), parentFkeyProp.getValue());
	}
	
}
