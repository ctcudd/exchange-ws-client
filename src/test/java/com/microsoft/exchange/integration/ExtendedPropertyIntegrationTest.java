/**
 * 
 */
package com.microsoft.exchange.integration;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.UUID;

import org.apache.commons.lang.StringUtils;
import org.junit.Test;
import org.springframework.dao.support.DataAccessUtils;

import com.microsoft.exchange.messages.UpdateItem;
import com.microsoft.exchange.messages.UpdateItemResponse;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.MapiPropertyTypeType;
import com.microsoft.exchange.types.NonEmptyArrayOfItemChangesType;
import com.microsoft.exchange.types.NonEmptyArrayOfPropertyValuesType;
import com.microsoft.exchange.types.PathToExtendedFieldType;
import com.microsoft.exchange.types.SetItemFieldType;


/**
 * @author ctcudd
 *
 */
public class ExtendedPropertyIntegrationTest extends BaseExchangeCalendarDataDaoIntegrationTest{

	private static final UUID seed = UUID.randomUUID();
	
	@Test
	public void isAutowired(){
		assertNotNull(exchangeCalendarDataDao.getRequestFactory());
		assertNotNull(upn);
	}
	
	protected PathToExtendedFieldType getExtendedFieldType(String propertyName, MapiPropertyTypeType type){
		PathToExtendedFieldType extendedFieldType = new PathToExtendedFieldType();
		extendedFieldType.setPropertyName(propertyName);
		extendedFieldType.setPropertySetId(seed.toString());
		extendedFieldType.setPropertyType(type);
		return extendedFieldType;
	}
	/**
	 * Control test showing the ability to create and retrieve an item on the exchange server with a single extended property
	 */
	@Test
	public void createAndGetSingleStringExtendedProperty(){
		
		//create extended property path
		PathToExtendedFieldType extendedFieldType = getExtendedFieldType("propertyName", MapiPropertyTypeType.STRING);
		
		//config request factory to return the extended property
		exchangeCalendarDataDao.getRequestFactory().setItemExtendedPropertyPaths(Collections.singleton(extendedFieldType));
		
		//create the calendar item
		CalendarItemType c = new CalendarItemType();
		
		//create extended property
		ExtendedPropertyType expectedProperty = new ExtendedPropertyType();
		
		//add previously created property path 
		expectedProperty.setExtendedFieldURI(extendedFieldType);
		
		//add prop value
		expectedProperty.setValue("qwertyuiopasdfghjklzxcvbnm");
		
		//add the extended prop to the calendar item
		c.getExtendedProperties().add(expectedProperty);
		
		//create the item on the exchange server
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(calendarItems);
		assertEquals(1, calendarItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(calendarItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		
		//make sure a single extended property is returned
		List<ExtendedPropertyType> extendedProperties = createdItem.getExtendedProperties();
		assertNotNull(extendedProperties);
		assertEquals(1, extendedProperties.size());
		ExtendedPropertyType createdProperty = DataAccessUtils.singleResult(extendedProperties);
		
		//ensure the property is equal
		assertEquals(expectedProperty, createdProperty);
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, expectedItemId);
		assertTrue(deleteSuccess);
	}
	
	/**
	 * Test showing that it is not possible to add two extended properties with the same path.  The second property will overwrite the second.
	 */
	@Test
	public void createAndGetDuplicateExtendedProperties(){
		//create extended property path
		PathToExtendedFieldType extendedFieldType = getExtendedFieldType("propertyName", MapiPropertyTypeType.STRING);
		
		//config request factory to return the extended property
		exchangeCalendarDataDao.getRequestFactory().setItemExtendedPropertyPaths(Collections.singleton(extendedFieldType));
		
		//create the calendar item
		CalendarItemType c = new CalendarItemType();
		
		//create extended property
		ExtendedPropertyType expectedProperty = new ExtendedPropertyType();
		
		//add previously created property path 
		expectedProperty.setExtendedFieldURI(extendedFieldType);
		
		//create second extended property with same path
		ExtendedPropertyType expectedProperty2 = expectedProperty;
		
		//add prop values
		expectedProperty.setValue("qwertyuiopasdfghjklzxcvbnm");
		expectedProperty2.setValue("ONEOFTHESETHINGSISNOTLIKETHEOTHER");
		
		//add the extended prop to the calendar item
		c.getExtendedProperties().add(expectedProperty);
		c.getExtendedProperties().add(expectedProperty2);
		
		//create the item on the exchange server
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(calendarItems);
		assertEquals(1, calendarItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(calendarItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		
		//make sure a single extended property is returned
		List<ExtendedPropertyType> extendedProperties = createdItem.getExtendedProperties();
		assertNotNull(extendedProperties);
		assertEquals(1, extendedProperties.size());
		ExtendedPropertyType createdProperty = DataAccessUtils.singleResult(extendedProperties);
		
		//ensure the property is equal
		assertEquals(expectedProperty2, createdProperty);
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, expectedItemId);
		assertTrue(deleteSuccess);
	}
	
	/**
	 * The following test shows that two extend properties with the same propertysetid but different property names can be stored/retrieved
	 */
	@Test
	public void createAndGetDuplicatePropertySetIdExtendedProperties(){
		PathToExtendedFieldType extendedFieldType1 = getExtendedFieldType("propertyName1",MapiPropertyTypeType.STRING);
		PathToExtendedFieldType extendedFieldType2 = getExtendedFieldType("propertyName2",MapiPropertyTypeType.STRING);
		
		Set<PathToExtendedFieldType> paths = new HashSet<PathToExtendedFieldType>();
		
		/**
		 * Attempts to retrieve all properties by propertysetid... This fails as
		 * either propertyName or propertyId, but not both, must be set
		 * otherwise 'The extended property attribute combination is invalid'
		 * 
			PathToExtendedFieldType partialPath = new PathToExtendedFieldType();
			partialPath.setPropertySetId(extendedFieldType2.getPropertySetId());
			partialPath.setPropertyType(extendedFieldType2.getPropertyType());
			paths.add(partialPath);
		*/
		
		paths.add(extendedFieldType1);
		paths.add(extendedFieldType2);
		
		//config request factory to return the extended property
		exchangeCalendarDataDao.getRequestFactory().setItemExtendedPropertyPaths(paths);
		
		//create the calendar item
		CalendarItemType c = new CalendarItemType();
		
		//create extended property
		ExtendedPropertyType expectedProperty1 = new ExtendedPropertyType();
		ExtendedPropertyType expectedProperty2 = new ExtendedPropertyType();
		
		//add property path 
		expectedProperty1.setExtendedFieldURI(extendedFieldType1);
		expectedProperty2.setExtendedFieldURI(extendedFieldType2);
		
		//add prop value
		expectedProperty1.setValue("qwertyuiopasdfghjklzxcvbnm");
		expectedProperty2.setValue("ONEOFTHESETHINGSISNOTLIKETHEOTHER");
		
		//add the extended prop to the calendar item
		c.getExtendedProperties().add(expectedProperty1);
		c.getExtendedProperties().add(expectedProperty2);
		
		//create the item on the exchange server
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(calendarItems);
		assertEquals(1, calendarItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(calendarItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		
		//make sure two extended properties are returned
		List<ExtendedPropertyType> extendedProperties = createdItem.getExtendedProperties();
		assertNotNull(extendedProperties);
		assertEquals(2, extendedProperties.size());
		
		//ensure both properties were retrieved
		assertTrue(extendedProperties.contains(expectedProperty1));
		assertTrue(extendedProperties.contains(expectedProperty2));
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, expectedItemId);
		assertTrue(deleteSuccess);
	}
	
	@Test
	public void createAndGetItemWith2DimensionalStringArrayAsExtendedProperty(){
		PathToExtendedFieldType extendedFieldType = getExtendedFieldType("2DArray", MapiPropertyTypeType.STRING_ARRAY);

		exchangeCalendarDataDao.getRequestFactory().setItemExtendedPropertyPaths(java.util.Collections.singleton(extendedFieldType));

		//populate collection with addresses for 1st visitor
		Set<String> visitor1Emails = new HashSet<String>();
		visitor1Emails.add("ctcudd@wisc.edu");
		visitor1Emails.add("collin.cudd@wisc.edu");
		visitor1Emails.add("someaddress@me.com");
		
		//populate collection with addresses for 2nd visitor
		Set<String> visitor2Emails = new HashSet<String>();
		visitor2Emails.add("hamdrew@wisc.edu");
		visitor2Emails.add("hamathoy.drew@wisc.edu");
		
		Set<String> visitor3Emails = new HashSet<String>();
		visitor3Emails.add("billcosby@jello.com");
		visitor3Emails.add("bill@thecosbyshow.com");
		
		//construct a semi colon seperated list of strings for all visitor addresses
		Set<String> visitors = new HashSet<String>();
		visitors.add(StringUtils.join(visitor1Emails, ";"));
		visitors.add(StringUtils.join(visitor2Emails, ";"));
		
		//create the calendar item
		CalendarItemType c = new CalendarItemType();
		ExtendedPropertyType visitorsProperty = new ExtendedPropertyType();
		NonEmptyArrayOfPropertyValuesType values = new NonEmptyArrayOfPropertyValuesType();
		values.getValues().addAll(visitors);
		visitorsProperty.setValues(values);
		visitorsProperty.setExtendedFieldURI(extendedFieldType);
		c.getExtendedProperties().add(visitorsProperty);
		
		//create the item on the exchange server
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(calendarItems);
		assertEquals(1, calendarItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(calendarItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		
		//make sure the extended property is returned
		List<ExtendedPropertyType> extendedProperties = createdItem.getExtendedProperties();
		assertNotNull(extendedProperties);
		assertEquals(1, extendedProperties.size());
		
		//ensure both properties were retrieved
		ExtendedPropertyType retrievedProperty = DataAccessUtils.singleResult(extendedProperties);
		assertEquals(visitorsProperty, retrievedProperty);
		
		NonEmptyArrayOfPropertyValuesType retrievedValuesArray = retrievedProperty.getValues();
		List<String> retrievedValues = retrievedValuesArray.getValues();
		for(String s : retrievedValues){
			assertTrue(visitors.contains(s));
		}
		
		
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, expectedItemId);
		assertTrue(deleteSuccess);
	}

	@Test
	public void createAndUpdateAndGetItemWith2DimensionalStringArrayAsExtendedProperty(){
		PathToExtendedFieldType extendedFieldType = getExtendedFieldType("2DArray", MapiPropertyTypeType.STRING_ARRAY);

		exchangeCalendarDataDao.getRequestFactory().setItemExtendedPropertyPaths(java.util.Collections.singleton(extendedFieldType));

		//populate collection with addresses for 1st visitor
		Set<String> visitor1Emails = new HashSet<String>();
		visitor1Emails.add("ctcudd@wisc.edu");
		visitor1Emails.add("collin.cudd@wisc.edu");
		visitor1Emails.add("someaddress@me.com");
		
		//populate collection with addresses for 2nd visitor
		Set<String> visitor2Emails = new HashSet<String>();
		visitor2Emails.add("hamdrew@wisc.edu");
		visitor2Emails.add("hamathoy.drew@wisc.edu");
		
		Set<String> visitor3Emails = new HashSet<String>();
		visitor3Emails.add("billcosby@jello.com");
		visitor3Emails.add("bill@thecosbyshow.com");
		
		//construct a semi colon seperated list of strings for all visitor addresses
		Set<String> visitors = new HashSet<String>();
		visitors.add(StringUtils.join(visitor1Emails, ";"));
		visitors.add(StringUtils.join(visitor2Emails, ";"));
		
		//create the calendar item
		CalendarItemType c = new CalendarItemType();
		ExtendedPropertyType visitorsProperty = new ExtendedPropertyType();
		NonEmptyArrayOfPropertyValuesType values = new NonEmptyArrayOfPropertyValuesType();
		values.getValues().addAll(visitors);
		visitorsProperty.setValues(values);
		visitorsProperty.setExtendedFieldURI(extendedFieldType);
		c.getExtendedProperties().add(visitorsProperty);
		
		//create the item on the exchange server
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(calendarItems);
		assertEquals(1, calendarItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(calendarItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		
		//make sure the extended property is returned
		List<ExtendedPropertyType> extendedProperties = createdItem.getExtendedProperties();
		assertNotNull(extendedProperties);
		assertEquals(1, extendedProperties.size());
		
		//ensure both properties were retrieved
		ExtendedPropertyType retrievedProperty = DataAccessUtils.singleResult(extendedProperties);
		assertEquals(visitorsProperty, retrievedProperty);
		
		NonEmptyArrayOfPropertyValuesType retrievedValuesArray = retrievedProperty.getValues();
		List<String> retrievedValues = retrievedValuesArray.getValues();
		for(String s : retrievedValues){
			assertTrue(visitors.contains(s));
		}
		
		//now prepare to update the item
		//add subject
		String newSubject = "newSubject";
		createdItem.setSubject(newSubject);
		
		//remove old extended property?
		//createdItem.getExtendedProperties().remove(visitorsProperty);
		
		//construct new extended property
		visitorsProperty = new ExtendedPropertyType();
		//update values
		visitors.add(StringUtils.join(visitor3Emails,";"));
		values = new NonEmptyArrayOfPropertyValuesType();
		values.getValues().addAll(visitors);
		visitorsProperty.setValues(values);
		visitorsProperty.setExtendedFieldURI(extendedFieldType);

		//create setVisitors object
		SetItemFieldType setVisitorsProperty = exchangeCalendarDataDao.getRequestFactory().constructSetCalendarItemExtendedProperty(createdItem,visitorsProperty);
		
		//create setSubject field
		SetItemFieldType setSubject = exchangeCalendarDataDao.getRequestFactory().constructSetCalendarItemSubject(createdItem);
		
		//construct collection for changes
		Collection<SetItemFieldType> changeFields = new HashSet<SetItemFieldType>();
		changeFields.add(setSubject);
		changeFields.add(setVisitorsProperty);
		NonEmptyArrayOfItemChangesType changes = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItemChanges(createdItem,changeFields);
		
		UpdateItem request = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItem(createdItem, changes);
		UpdateItemResponse response = exchangeCalendarDataDao.getWebServices().updateItem(request);
		Set<ItemIdType> updatedItemIds = exchangeCalendarDataDao.getResponseUtils().parseUpdateItemResponse(response);
		ItemIdType updatedItemId = DataAccessUtils.singleResult(updatedItemIds);
		
		//retrieve the newly updated item
		calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, Collections.singleton(updatedItemId));
				
		//ensures a singular non null result
		assertNotNull(calendarItems);
		assertEquals(1, calendarItems.size());
		CalendarItemType updatedItem = DataAccessUtils.singleResult(calendarItems);
			
		//make sure the ids match
		assertEquals(updatedItemId, updatedItem.getItemId());
		
		//make sure the subject has been updated
		assertEquals(newSubject, updatedItem.getSubject());
		
		//make sure the extended property is returned
		extendedProperties = updatedItem.getExtendedProperties();
		assertNotNull(extendedProperties);
		assertEquals(1, extendedProperties.size());
		
		//ensure both properties were retrieved
		retrievedProperty = DataAccessUtils.singleResult(extendedProperties);
		assertEquals(visitorsProperty, retrievedProperty);
		
		retrievedValuesArray = retrievedProperty.getValues();
		retrievedValues = retrievedValuesArray.getValues();
		for(String s : retrievedValues){
			assertTrue(visitors.contains(s));
		}
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, expectedItemId);
		assertTrue(deleteSuccess);
	}
}
